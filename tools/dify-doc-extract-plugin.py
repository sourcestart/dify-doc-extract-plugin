import hashlib
import io
import re
import struct
from collections.abc import Generator
from typing import Any

import olefile
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File


class DifyDocExtractPluginTool(Tool):
    """Extract text and images from legacy .doc files without external programs."""

    _CLX_FC_INDEX = 33
    _BLIP_REC_TYPES = set(range(0xF018, 0xF030))

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        input_file = tool_parameters.get("input_file")
        if not isinstance(input_file, File):
            yield self.create_text_message("Error: Missing or invalid parameter 'input_file' (expected Dify file).")
            return

        filename = input_file.filename or "document.doc"
        extension = (input_file.extension or "").lower().lstrip(".")
        if extension != "doc" and not filename.lower().endswith(".doc"):
            yield self.create_text_message("Error: Invalid file format. Only .doc files are supported.")
            return

        blob = input_file.blob or b""
        if not blob:
            yield self.create_text_message("Error: Uploaded file is empty.")
            return

        try:
            result = self._extract_doc(blob)
        except ValueError as e:
            yield self.create_text_message(f"Error: {e}")
            return
        except Exception as e:
            yield self.create_text_message(f"Error: Failed to parse DOC file: {e}")
            return

        output_images = []
        base_name = re.sub(r"\.doc$", "", filename, flags=re.IGNORECASE) or "document"
        for idx, image in enumerate(result["images"], start=1):
            output_name = f"{base_name}_image_{idx}.{image['extension']}"
            output_images.append(
                {
                    "index": idx,
                    "filename": output_name,
                    "mime_type": image["mime_type"],
                    "size": len(image["data"]),
                    "source": image["source"],
                }
            )
            yield self.create_blob_message(
                blob=image["data"],
                meta={
                    "filename": output_name,
                    "file_name": output_name,
                    "mime_type": image["mime_type"],
                },
            )

        yield self.create_text_message(result["text"])

        yield self.create_json_message(
            {
                "status": "success",
                "source_file": filename,
                "text": result["text"],
                "text_length": len(result["text"]),
                "images": output_images,
                "image_count": len(output_images),
                "image_strategy": result["image_strategy"],
            }
        )

    def _extract_doc(self, blob: bytes) -> dict[str, Any]:
        if not olefile.isOleFile(io.BytesIO(blob)):
            raise ValueError("Input is not a valid OLE Compound File (.doc).")

        with olefile.OleFileIO(io.BytesIO(blob)) as ole:
            if not ole.exists("WordDocument"):
                raise ValueError("Invalid DOC: missing WordDocument stream.")

            word_stream = ole.openstream("WordDocument").read()
            fib = self._parse_fib(word_stream)

            table_name = "1Table" if fib["which_table"] else "0Table"
            if not ole.exists(table_name):
                alt = "0Table" if table_name == "1Table" else "1Table"
                if ole.exists(alt):
                    table_name = alt
                else:
                    raise ValueError("Invalid DOC: missing table stream (0Table/1Table).")

            table_stream = ole.openstream(table_name).read()
            text = self._extract_text_from_piece_table(word_stream, table_stream, fib)

            candidate_streams = {
                "WordDocument": word_stream,
                table_name: table_stream,
            }
            if ole.exists("Data"):
                candidate_streams["Data"] = ole.openstream("Data").read()

            images = self._extract_images(candidate_streams)

        return {
            "text": text,
            "images": images["items"],
            "image_strategy": images["strategy"],
        }

    def _parse_fib(self, word_stream: bytes) -> dict[str, int | bool]:
        if len(word_stream) < 64:
            raise ValueError("Invalid DOC: WordDocument stream is too short.")

        w_ident = struct.unpack_from("<H", word_stream, 0)[0]
        if w_ident != 0xA5EC:
            raise ValueError("Invalid DOC: unsupported Word binary header.")

        flags = struct.unpack_from("<H", word_stream, 0x0A)[0]
        which_table = bool(flags & 0x0200)

        csw = struct.unpack_from("<H", word_stream, 32)[0]
        cslw_offset = 34 + csw * 2
        if len(word_stream) < cslw_offset + 2:
            raise ValueError("Invalid DOC: corrupt FIB section.")

        cslw = struct.unpack_from("<H", word_stream, cslw_offset)[0]
        rglw_offset = cslw_offset + 2

        ccp_text = 0
        if cslw > 3 and len(word_stream) >= rglw_offset + (4 * cslw):
            ccp_text = struct.unpack_from("<I", word_stream, rglw_offset + 12)[0]

        cb_rg_fc_lcb_offset = rglw_offset + (4 * cslw)
        if len(word_stream) < cb_rg_fc_lcb_offset + 2:
            raise ValueError("Invalid DOC: missing FibRgFcLcb count.")

        cb_rg_fc_lcb = struct.unpack_from("<H", word_stream, cb_rg_fc_lcb_offset)[0]
        rg_fc_lcb_offset = cb_rg_fc_lcb_offset + 2

        fc_clx = 0
        lcb_clx = 0
        if cb_rg_fc_lcb > self._CLX_FC_INDEX:
            pair_offset = rg_fc_lcb_offset + self._CLX_FC_INDEX * 8
            if len(word_stream) >= pair_offset + 8:
                fc_clx, lcb_clx = struct.unpack_from("<II", word_stream, pair_offset)

        if lcb_clx == 0 and len(word_stream) >= 0x1AA:
            fc_clx = struct.unpack_from("<I", word_stream, 0x1A2)[0]
            lcb_clx = struct.unpack_from("<I", word_stream, 0x1A6)[0]

        if lcb_clx == 0:
            raise ValueError("Invalid DOC: Piece Table (CLX) not found.")

        return {
            "which_table": which_table,
            "fc_clx": fc_clx,
            "lcb_clx": lcb_clx,
            "ccp_text": ccp_text,
        }

    def _extract_text_from_piece_table(self, word_stream: bytes, table_stream: bytes, fib: dict[str, Any]) -> str:
        fc_clx = int(fib["fc_clx"])
        lcb_clx = int(fib["lcb_clx"])
        ccp_text = int(fib.get("ccp_text") or 0)

        if fc_clx < 0 or lcb_clx <= 0 or fc_clx + lcb_clx > len(table_stream):
            raise ValueError("Invalid DOC: CLX range out of table stream bounds.")

        clx = table_stream[fc_clx : fc_clx + lcb_clx]
        piece_table = self._find_piece_table(clx)
        if not piece_table:
            raise ValueError("Invalid DOC: Piece Table payload not found in CLX.")

        cp_array, pcd_array = piece_table
        text_parts: list[str] = []

        for i in range(len(pcd_array)):
            cp_start = cp_array[i]
            cp_end = cp_array[i + 1]

            if cp_end <= cp_start:
                continue

            if ccp_text and cp_start >= ccp_text:
                break

            char_count = cp_end - cp_start
            if ccp_text and cp_end > ccp_text:
                char_count = max(0, ccp_text - cp_start)
            if char_count <= 0:
                continue

            pcd = pcd_array[i]
            fc_raw = struct.unpack_from("<I", pcd, 2)[0]
            compressed = bool(fc_raw & 0x40000000)
            fc_value = fc_raw & 0x3FFFFFFF

            if compressed:
                byte_offset = fc_value // 2
                byte_count = char_count
                piece_bytes = word_stream[byte_offset : byte_offset + byte_count]
                text_parts.append(piece_bytes.decode("cp1252", errors="replace"))
            else:
                byte_offset = fc_value
                byte_count = char_count * 2
                piece_bytes = word_stream[byte_offset : byte_offset + byte_count]
                text_parts.append(piece_bytes.decode("utf-16le", errors="replace"))

        text = "".join(text_parts)
        text = text.replace("\r", "\n")
        text = re.sub(r"\x07", "\t", text)
        text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def _find_piece_table(self, clx: bytes) -> tuple[list[int], list[bytes]] | None:
        i = 0
        length = len(clx)

        while i < length:
            clxt = clx[i]
            if clxt == 0x01:
                if i + 3 > length:
                    break
                cb_grpprl = struct.unpack_from("<H", clx, i + 1)[0]
                i += 3 + cb_grpprl
                continue

            if clxt == 0x02:
                if i + 5 > length:
                    break
                lcb = struct.unpack_from("<I", clx, i + 1)[0]
                start = i + 5
                end = start + lcb
                if end > length:
                    break

                plc_pcd = clx[start:end]
                if len(plc_pcd) < 4:
                    break

                pcd_count = (len(plc_pcd) - 4) // 12
                if pcd_count <= 0:
                    break

                cp_array = [
                    struct.unpack_from("<I", plc_pcd, j * 4)[0]
                    for j in range(pcd_count + 1)
                ]
                pcd_offset = (pcd_count + 1) * 4
                pcd_array = [
                    plc_pcd[pcd_offset + j * 8 : pcd_offset + (j + 1) * 8]
                    for j in range(pcd_count)
                ]
                return cp_array, pcd_array

            i += 1

        return None

    def _extract_images(self, streams: dict[str, bytes]) -> dict[str, Any]:
        seen: set[str] = set()
        images: list[dict[str, Any]] = []

        for stream_name, stream_data in streams.items():
            for image in self._extract_officeart_blips(stream_data):
                digest = hashlib.sha1(image["data"]).hexdigest()
                if digest in seen:
                    continue
                seen.add(digest)
                image["source"] = f"officeart:{stream_name}"
                images.append(image)

        if images:
            return {"items": images, "strategy": "officeart_blip"}

        for stream_name, stream_data in streams.items():
            for image in self._signature_scan_images(stream_data):
                digest = hashlib.sha1(image["data"]).hexdigest()
                if digest in seen:
                    continue
                seen.add(digest)
                image["source"] = f"signature_fallback:{stream_name}"
                images.append(image)

        return {"items": images, "strategy": "signature_fallback" if images else "none"}

    def _extract_officeart_blips(self, data: bytes) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        i = 0
        size = len(data)

        while i + 8 <= size:
            rec_ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, i)
            if rec_type in self._BLIP_REC_TYPES and rec_len > 0 and i + 8 + rec_len <= size:
                payload = data[i + 8 : i + 8 + rec_len]
                parsed = self._extract_image_from_payload(payload, rec_type)
                if parsed:
                    out.append(parsed)
                i += 8 + rec_len
                continue
            i += 1

        return out

    def _extract_image_from_payload(self, payload: bytes, rec_type: int) -> dict[str, Any] | None:
        magic_image = self._find_single_image(payload)
        if magic_image:
            return magic_image

        if rec_type in {0xF01E, 0xF01F, 0xF029}:
            dib = self._extract_dib(payload)
            if dib:
                return {
                    "data": dib,
                    "mime_type": "image/bmp",
                    "extension": "bmp",
                }

        return None

    def _find_single_image(self, payload: bytes) -> dict[str, Any] | None:
        jpeg = self._carve_jpeg(payload)
        if jpeg:
            return {"data": jpeg, "mime_type": "image/jpeg", "extension": "jpg"}

        png = self._carve_png(payload)
        if png:
            return {"data": png, "mime_type": "image/png", "extension": "png"}

        gif = self._carve_gif(payload)
        if gif:
            return {"data": gif, "mime_type": "image/gif", "extension": "gif"}

        bmp = self._carve_bmp(payload)
        if bmp:
            return {"data": bmp, "mime_type": "image/bmp", "extension": "bmp"}

        tiff = self._carve_tiff(payload)
        if tiff:
            return {"data": tiff, "mime_type": "image/tiff", "extension": "tiff"}

        return None

    def _signature_scan_images(self, data: bytes) -> list[dict[str, Any]]:
        images: list[dict[str, Any]] = []
        scans = [
            (self._carve_all_jpegs, "image/jpeg", "jpg"),
            (self._carve_all_pngs, "image/png", "png"),
            (self._carve_all_gifs, "image/gif", "gif"),
            (self._carve_all_bmps, "image/bmp", "bmp"),
            (self._carve_all_tiffs, "image/tiff", "tiff"),
        ]

        for carver, mime, ext in scans:
            for img in carver(data):
                images.append({"data": img, "mime_type": mime, "extension": ext})

        return images

    def _carve_all_jpegs(self, data: bytes) -> list[bytes]:
        out: list[bytes] = []
        cursor = 0
        while True:
            start = data.find(b"\xFF\xD8\xFF", cursor)
            if start < 0:
                break
            end = data.find(b"\xFF\xD9", start + 3)
            if end < 0:
                break
            out.append(data[start : end + 2])
            cursor = end + 2
        return out

    def _carve_all_pngs(self, data: bytes) -> list[bytes]:
        out: list[bytes] = []
        sig = b"\x89PNG\r\n\x1a\n"
        cursor = 0
        while True:
            start = data.find(sig, cursor)
            if start < 0:
                break
            iend = data.find(b"IEND", start + 8)
            if iend < 0 or iend + 8 > len(data):
                break
            out.append(data[start : iend + 8])
            cursor = iend + 8
        return out

    def _carve_all_gifs(self, data: bytes) -> list[bytes]:
        out: list[bytes] = []
        cursor = 0
        while True:
            start = data.find(b"GIF8", cursor)
            if start < 0:
                break
            end = data.find(b"\x3B", start + 6)
            if end < 0:
                break
            out.append(data[start : end + 1])
            cursor = end + 1
        return out

    def _carve_all_bmps(self, data: bytes) -> list[bytes]:
        out: list[bytes] = []
        cursor = 0
        while True:
            start = data.find(b"BM", cursor)
            if start < 0 or start + 6 > len(data):
                break
            size = struct.unpack_from("<I", data, start + 2)[0]
            if size > 54 and start + size <= len(data):
                out.append(data[start : start + size])
                cursor = start + size
            else:
                cursor = start + 2
        return out

    def _carve_all_tiffs(self, data: bytes) -> list[bytes]:
        out: list[bytes] = []
        for sig in (b"II*\x00", b"MM\x00*"):
            cursor = 0
            while True:
                start = data.find(sig, cursor)
                if start < 0:
                    break
                out.append(data[start:])
                cursor = start + 4
        return out

    def _carve_jpeg(self, data: bytes) -> bytes | None:
        start = data.find(b"\xFF\xD8\xFF")
        if start < 0:
            return None
        end = data.find(b"\xFF\xD9", start + 3)
        if end < 0:
            return None
        return data[start : end + 2]

    def _carve_png(self, data: bytes) -> bytes | None:
        sig = b"\x89PNG\r\n\x1a\n"
        start = data.find(sig)
        if start < 0:
            return None
        iend = data.find(b"IEND", start + 8)
        if iend < 0 or iend + 8 > len(data):
            return None
        return data[start : iend + 8]

    def _carve_gif(self, data: bytes) -> bytes | None:
        start = data.find(b"GIF8")
        if start < 0:
            return None
        end = data.find(b"\x3B", start + 6)
        if end < 0:
            return None
        return data[start : end + 1]

    def _carve_bmp(self, data: bytes) -> bytes | None:
        start = data.find(b"BM")
        if start < 0 or start + 6 > len(data):
            return None
        size = struct.unpack_from("<I", data, start + 2)[0]
        if size <= 54 or start + size > len(data):
            return None
        return data[start : start + size]

    def _carve_tiff(self, data: bytes) -> bytes | None:
        for sig in (b"II*\x00", b"MM\x00*"):
            start = data.find(sig)
            if start >= 0:
                return data[start:]
        return None

    def _extract_dib(self, payload: bytes) -> bytes | None:
        header_sizes = {12, 40, 52, 56, 108, 124}
        for offset in (0, 16, 17, 32, 33):
            if offset + 4 > len(payload):
                continue
            dib_header_size = struct.unpack_from("<I", payload, offset)[0]
            if dib_header_size not in header_sizes:
                continue

            dib = payload[offset:]
            if len(dib) < dib_header_size + 4:
                continue

            bf_type = b"BM"
            bf_size = 14 + len(dib)
            bf_reserved = 0
            bf_off_bits = 14 + dib_header_size
            bmp_header = struct.pack("<2sIHI", bf_type, bf_size, bf_reserved, bf_off_bits)
            return bmp_header + dib

        return None

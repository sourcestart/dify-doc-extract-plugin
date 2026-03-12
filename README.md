# dify-doc-extract-plugin

DOC extraction tool plugin for Dify.

## Version

- Current version: `0.0.1`
- Capability scope: extract `.doc` text and images

## Features

- Parse legacy Microsoft Word `.doc` files directly in Python (no external programs)
- Stage 1 text extraction:
  - OLE stream parsing (`WordDocument`, `0Table` / `1Table`)
  - Piece Table (`CLX` / `PlcPcd`) based text reconstruction
- Stage 2 image extraction:
  - OfficeArt/Blip record parsing as primary strategy
  - Signature scan fallback only when OfficeArt extraction finds no image

## Tool

- Tool name: `doc_extractor`
- Input:
  - `input_file` (`file`, required): must be a `.doc` file
- Output:
  - JSON message with:
    - `text`, `text_length`
    - `images` metadata list
    - `image_count`, `image_strategy`
  - Blob messages for each extracted image

## Dependencies

- Python dependencies:
  - `dify_plugin>=0.4.0,<0.7.0`
  - `olefile>=0.47,<1.0`
- External runtime dependency:
  - None

## Error Handling

Explicit errors are returned for:

- Missing or invalid uploaded file parameter
- Uploaded file is not `.doc`
- Uploaded file is empty
- Input is not a valid OLE Compound File
- Required DOC streams are missing (`WordDocument`, table stream)
- Piece Table (`CLX`) missing/corrupt
- Unexpected parsing failures with clear message prefix

## Temporary Files

- This plugin parses in-memory bytes only.
- No temporary files are created, so no residual temp artifacts are left.

## Limits

- This version targets classic binary `.doc` (Word 97-2003 family).
- Some uncommon OfficeArt variants may still fall back to signature scan behavior.

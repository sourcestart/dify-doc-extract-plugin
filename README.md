# DOC Extract Plugin for Dify

A Dify plugin for extracting text and image content from Microsoft Word `.doc` files.

## Version Information

- Current version: `0.0.1`
- Capability scope: Extract text and images from `.doc` files

## Feature Characteristics

- Directly parse legacy Microsoft Word `.doc` files using Python (no external programs required)
- Stage 1 text extraction:
  - OLE stream parsing (`WordDocument`, `0Table` / `1Table`)
  - Text reconstruction based on Piece Table (`CLX` / `PlcPcd`)
- Stage 2 image extraction:
  - OfficeArt/Blip record parsing as the primary strategy
  - Signature scan fallback only when OfficeArt extraction finds no images

## Tool Details

- Tool name: `doc_extractor`
- Input parameters:
  - `input_file` (`file`, required): Must be a `.doc` file
- Output results:
  - Blob messages for each extracted image (if images exist)
  - Text message containing the extracted text
  - JSON message containing processing results with the following fields:
    - `status`: Processing status ("success")
    - `source_file`: Source filename
    - `text`: Extracted text content
    - `text_length`: Length of text
    - `images`: Image metadata list
    - `image_count`: Number of images
    - `image_strategy`: Image extraction strategy

## Installation Dependencies

- Python dependencies:
  - `dify_plugin>=0.4.0,<0.7.0`
  - `olefile>=0.47,<1.0`
- External runtime dependencies:
  - None

## Usage Instructions

1. Enable this plugin in the Dify application
2. Select the `DOC Extractor` tool
3. Upload the `.doc` file to be processed
4. The plugin will return the extracted text content and image files

## Error Handling

Clear errors are returned for the following situations:

- Missing or invalid upload file parameter
- Uploaded file is not in `.doc` format
- Uploaded file is empty
- Input is not a valid OLE compound file
- Required DOC streams are missing (`WordDocument`, table stream)
- Piece Table (`CLX`) missing/corrupted
- Unexpected parsing failure (with clear message prefix)

## Technical Details

This plugin uses the olefile library to directly parse the internal structure of Word documents without installing Microsoft Word or LibreOffice. It can reliably extract text and embedded images from legacy `.doc` files.

## Author

Created by zhanghong

## License

MIT License

## Temporary Files

- This plugin parses in-memory bytes only.
- No temporary files are created, so no residual temp artifacts are left.

## Limits

- This version targets classic binary `.doc` (Word 97-2003 family).
- Some uncommon OfficeArt variants may still fall back to signature scan behavior.

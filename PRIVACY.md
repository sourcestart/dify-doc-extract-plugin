# Privacy Policy

## Data Processing

The DOC Extract Plugin follows strict data privacy protection principles when processing documents:

1. **Data does not leave user environment**: The plugin processes uploaded `.doc` files only within the Dify platform and does not send file content to any external servers.

2. **Temporary file management**: Temporary files generated during processing will be immediately deleted upon completion of operations, leaving no persistent storage in the system.

3. **Content extraction**: The plugin extracts only text and image content from documents and returns this content as processing results to the caller, without collecting any additional metadata or user information.

## Data Security

1. **No data collection**: The plugin does not collect, store, or transmit users' document content or usage habits.

2. **Memory-based processing**: All document parsing occurs in memory, with original files not written to disk.

3. **Secure processing**: The plugin uses secure document parsing libraries to avoid common file processing vulnerabilities.

## User Control

1. **Complete transparency**: Users have full control over uploaded documents and the resulting processed outputs.

2. **Local processing**: The entire document extraction process is completed in the local environment, ensuring the security of sensitive information.

## Third-party Dependencies

This plugin uses the following third-party libraries, all of which have been confirmed to meet security standards:

- `olefile`: Used to parse OLE compound document formats
- `dify_plugin`: Dify official plugin framework

For further information about the data processing methods of this plugin, please contact the developer.

import 'dart:typed_data';

const bool supportsFileAccess = false;

Future<Uint8List?> readFileBytes(String path) async => null;

Uint8List? readFileBytesSync(String path) => null;

Future<void> writeFileBytes(String path, Uint8List bytes) async {
  throw UnsupportedError(
    'Filesystem path access is not supported on this platform. Pass bytes instead.',
  );
}

void writeFileBytesSync(String path, Uint8List bytes) {
  throw UnsupportedError(
    'Filesystem path access is not supported on this platform. Pass bytes instead.',
  );
}

bool isDirectoryPathSync(String path) => false;

bool isFilePathSync(String path) => false;

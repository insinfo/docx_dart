import 'dart:typed_data';

import 'file_access_stub.dart' if (dart.library.io) 'file_access_io.dart' as impl;

bool get supportsFileAccess => impl.supportsFileAccess;

Future<Uint8List?> readFileBytes(String path) => impl.readFileBytes(path);

Uint8List? readFileBytesSync(String path) => impl.readFileBytesSync(path);

Future<void> writeFileBytes(String path, Uint8List bytes) =>
    impl.writeFileBytes(path, bytes);

void writeFileBytesSync(String path, Uint8List bytes) =>
    impl.writeFileBytesSync(path, bytes);

bool isDirectoryPathSync(String path) => impl.isDirectoryPathSync(path);

bool isFilePathSync(String path) => impl.isFilePathSync(path);

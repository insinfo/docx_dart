import 'dart:io';
import 'dart:typed_data';

const bool supportsFileAccess = true;

Future<Uint8List?> readFileBytes(String path) async {
  final file = File(path);
  if (!await file.exists()) {
    return null;
  }
  return Uint8List.fromList(await file.readAsBytes());
}

Uint8List? readFileBytesSync(String path) {
  final file = File(path);
  if (!file.existsSync()) {
    return null;
  }
  return Uint8List.fromList(file.readAsBytesSync());
}

Future<void> writeFileBytes(String path, Uint8List bytes) async {
  await File(path).writeAsBytes(bytes, flush: true);
}

void writeFileBytesSync(String path, Uint8List bytes) {
  File(path).writeAsBytesSync(bytes, flush: true);
}

bool isDirectoryPathSync(String path) => FileSystemEntity.isDirectorySync(path);

bool isFilePathSync(String path) => FileSystemEntity.isFileSync(path);

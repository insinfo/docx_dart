import 'dart:io';

String testFile(String filename) {
  final relative = 'test/test_files/$filename';
  final file = File(relative);
  if (!file.existsSync()) {
    throw StateError('Expected test file at ${file.path}');
  }
  return file.path;
}

// docx/image/exceptions.dart
class InvalidImageStreamError implements Exception {
  final String message;
  InvalidImageStreamError(this.message);
  @override String toString() => message;
}

class UnexpectedEndOfFileError implements Exception {
  final String message;
  UnexpectedEndOfFileError(this.message);
  @override String toString() => message;
}

class UnrecognizedImageError implements Exception {
  final String message;
  UnrecognizedImageError(this.message);
  @override String toString() => message;
}
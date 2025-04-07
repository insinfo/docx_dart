/// Path: lib/src/exceptions.dart
/// Based on python-docx: docx/exceptions.py

/// Generic base error class for python_docx_dart errors.
class PythonDocxError implements Exception {
  final String message;

  /// Creates a [PythonDocxError] with the given [message].
  PythonDocxError(this.message);

  @override
  String toString() => 'PythonDocxError: $message';
}

/// Raised when an invalid merge region is specified in a request to merge table cells.
class InvalidSpanError extends PythonDocxError {
  /// Creates an [InvalidSpanError] with the given [message].
  InvalidSpanError(super.message);
}

/// Raised when invalid XML is encountered, such as on attempt to access a missing
/// required child element.
class InvalidXmlError extends PythonDocxError {
  /// Creates an [InvalidXmlError] with the given [message].
  InvalidXmlError(super.message);
}
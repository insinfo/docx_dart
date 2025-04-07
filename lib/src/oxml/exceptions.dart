/// Path: lib/src/oxml/exceptions.dart
/// Based on python-docx: docx/oxml/exceptions.py
///
/// Exceptions specific to the oxml sub-package, dealing with Open XML
/// element representation and manipulation.

import '../exceptions.dart'; // Import the base PythonDocxError

/// Generic error class for oxml operations.
class XmlchemyError extends PythonDocxError {
  /// Creates an [XmlchemyError] with the given [message].
  XmlchemyError(super.message);

  @override
  String toString() => 'XmlchemyError: $message';
}

/// Raised when invalid XML is encountered during oxml processing,
/// such as attempting to access a missing required child element.
///
/// Note: This inherits from [XmlchemyError], establishing a hierarchy where
/// oxml-specific XML errors are distinct from more general `PythonDocxError`s
/// or potentially other `InvalidXmlError` contexts if they were defined elsewhere.
class InvalidXmlError extends XmlchemyError {
  /// Creates an [InvalidXmlError] with the given [message].
  InvalidXmlError(super.message);

  // Inherits toString from XmlchemyError, which includes the class name prefix.
}
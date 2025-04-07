/// Path: lib/src/opc/exceptions.dart
/// Based on python-docx: docx/opc/exceptions.py
///
/// Exceptions specific to the OPC (Open Packaging Conventions) handling.

/// Base error class for OPC-related errors.
class OpcError implements Exception {
  final String message;

  /// Creates an [OpcError] with the given [message].
  OpcError(this.message);

  @override
  String toString() => 'OpcError: $message';
}

/// Raised when a package cannot be found at the specified path.
class PackageNotFoundError extends OpcError {
  /// Creates a [PackageNotFoundError] indicating the package at [path] was not found.
  PackageNotFoundError(String path) : super("Package not found at '$path'");

  /// Creates a [PackageNotFoundError] with a custom [message].
  PackageNotFoundError.withMessage(super.message);
}
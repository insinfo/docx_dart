/// Path: lib/src/opc/packuri.dart
/// Based on python-docx: docx/opc/packuri.py
///
/// Provides the PackUri value type for representing internal OPC package URIs.

import 'package:path/path.dart' as p;

/// Represents a URI within an OPC package, always starting with '/'.
/// Provides access to pack URI components like base URI and filename.
/// Immutable.
class PackUri {
  /// The validated, normalized URI string (e.g., "/word/document.xml").
  final String uri;

  /// Creates a [PackUri] instance from a URI string.
  ///
  /// Throws [ArgumentError] if [uriString] does not start with '/'.
  PackUri(String uriString) : uri = _validateAndNormalize(uriString);

  /// Validates that the URI starts with '/' and normalizes it.
  static String _validateAndNormalize(String uriString) {
    if (!uriString.startsWith('/')) {
      throw ArgumentError("PackURI must begin with slash, got '$uriString'");
    }
    // Use url context for POSIX-style normalization
    return p.url.normalize(uriString);
  }

  /// Creates the absolute [PackUri] by resolving [relativeRef] against [baseUri].
  ///
  /// [baseUri] must be an absolute Pack URI string (starting with '/').
  /// [relativeRef] is the relative path string.
  static PackUri fromRelativeRef(String baseUri, String relativeRef) {
    if (!baseUri.startsWith('/')) {
      throw ArgumentError("baseUri must begin with slash, got '$baseUri'");
    }
    // Ensure baseUri ends with '/' if it's a directory-like structure
    // p.url.join handles this reasonably well, but explicit check can be safer
    final String normalizedBase = baseUri == '/' ? '/' : p.url.dirname(baseUri);
    final String joinedUri = p.url.join(normalizedBase, relativeRef);
    // Normalize resolves '..' and ensures it starts with '/'
    final String absoluteUri = p.url.normalize(joinedUri);
    return PackUri(absoluteUri);
  }

  /// The base URI of this pack URI (the directory portion).
  ///
  /// Example: `'/ppt/slides'` for `'/ppt/slides/slide1.xml'`.
  /// Returns `'/'` for the root URI `'/'`.
  String get baseUri {
    // p.url.dirname('/') returns '.', which is not desired for OPC base URI.
    if (uri == '/') return '/';
    return p.url.dirname(uri);
  }

  /// The extension portion of this pack URI without the leading dot.
  ///
  /// Example: `'xml'` for `'/word/document.xml'`. Returns an empty string if
  /// there is no extension.
  String get ext {
    final extension = p.url.extension(uri);
    return extension.startsWith('.') ? extension.substring(1) : extension;
  }

  /// The filename portion of this pack URI.
  ///
  /// Example: `'slide1.xml'` for `'/ppt/slides/slide1.xml'`.
  /// Returns `''` for the root URI `'/'`.
  String get filename => p.url.basename(uri);


  /// The numeric index parsed from the filename, or `null`.
  ///
  /// Assumes filenames like "item1.ext", "image42.png".
  /// Returns `null` if the filename doesn't end with digits before the extension
  /// or if the URI is '/'.
  /// Example: `21` for `'/ppt/slides/slide21.xml'`, `null` for `'/ppt/presentation.xml'`.
  int? get idx {
    final namePart = p.url.basenameWithoutExtension(uri);
    if (namePart.isEmpty) return null;
    // Regex to capture potential trailing digits after some letters
    final match = RegExp(r'([a-zA-Z]+)([1-9]\d*)$').firstMatch(namePart);
     if (match == null || match.group(2) == null) return null;
     return int.tryParse(match.group(2)!);
  }


  /// The pack URI string without the leading slash.
  ///
  /// This form is typically used as the member name in a Zip archive.
  /// Returns `''` for the root URI `'/'`.
  String get membername => uri == '/' ? '' : uri.substring(1);

  /// Returns the relative path string from [baseUriPath] to this URI.
  ///
  /// [baseUriPath] must be an absolute Pack URI string (starting with '/').
  /// Example: `PackUri('/ppt/slideLayouts/slideLayout1.xml').relativeRef('/ppt/slides')`
  /// returns `'../slideLayouts/slideLayout1.xml'`.
  String relativeRef(String baseUriPath) {
     if (!baseUriPath.startsWith('/')) {
      throw ArgumentError("baseUriPath must begin with slash, got '$baseUriPath'");
    }
    // p.relative has issues when `from` is '/', handle explicitly
    if (baseUriPath == '/') return membername;
    return p.url.relative(uri, from: baseUriPath);
  }


  /// The [PackUri] of the relationship part corresponding to this URI.
  ///
  /// Example: `'/word/_rels/document.xml.rels'` for `'/word/document.xml'`.
  PackUri get relsUri {
     final name = filename;
     // Handle root case explicitly
     final relsFilename = name.isEmpty ? '.rels' : '$name.rels';
     final base = baseUri;
     // Use p.url.join which handles '/' correctly
     final relsUriString = p.url.join(base, '_rels', relsFilename);
     // Normalize to ensure clean path starting with '/'
     return PackUri(p.url.normalize(relsUriString));
  }


  @override
  String toString() => uri;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is PackUri && runtimeType == other.runtimeType && uri == other.uri;

  @override
  int get hashCode => uri.hashCode;

  // --- Static known URIs ---

  /// Represents the root of the OPC package ('/').
  static final PackUri packageUri = PackUri('/');

  /// Represents the Content Types part URI ('/[Content_Types].xml').
  static final PackUri contentTypesUri = PackUri('/[Content_Types].xml');
}

// --- Top-level constants for convenience ---

/// Represents the root of the OPC package ('/').
final PackUri PACKAGE_URI = PackUri.packageUri;

/// Represents the Content Types part URI ('/[Content_Types].xml').
final PackUri CONTENT_TYPES_URI = PackUri.contentTypesUri;
/// Path: lib/src/oxml/ns.dart
/// Based on python-docx: docx/oxml/ns.py
///
/// Namespace-related objects and utility functions for Open XML.

/// Map of namespace prefixes to their corresponding URIs.
const Map<String, String> nsmap = {
  "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  "dc": "http://purl.org/dc/elements/1.1/",
  "dcmitype": "http://purl.org/dc/dcmitype/",
  "dcterms": "http://purl.org/dc/terms/",
  "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
  "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
  "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
  "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "xml": "http://www.w3.org/XML/1998/namespace",
  "xsi": "http://www.w3.org/2001/XMLSchema-instance",
};

/// Reverse map of namespace URIs to their corresponding prefixes.
final Map<String, String> pfxmap =
    nsmap.map((key, value) => MapEntry(value, key));

/// Represents an XML tag with a namespace prefix (e.g., "w:p").
/// Provides easy conversion to Clark notation ("{uri}localPart") and access
/// to individual components. Immutable.
class NamespacePrefixedTag {
  /// The original prefixed tag string (e.g., "w:p").
  final String nstag;

  /// The namespace prefix (e.g., "w").
  final String nspfx;

  /// The local part of the tag (e.g., "p").
  final String localPart;

  /// The namespace URI associated with the prefix (e.g., "http://.../main").
  final String nsuri;

  /// Creates a [NamespacePrefixedTag] from a string like "pfx:localPart".
  /// Throws [ArgumentError] if the format is incorrect or the prefix is unknown.
  NamespacePrefixedTag(String tag)
      // Use initializer list to calculate final fields based on the constructor argument 'tag'
      : nstag = tag, // Store original tag
        nspfx = _extractPrefix(tag),
        localPart = _extractLocalPart(tag),
        nsuri = _lookupUri(_extractPrefix(tag)); // Recalculate prefix for lookup

  /// Creates a [NamespacePrefixedTag] from a Clark notation string like "{uri}localPart".
  /// Throws [ArgumentError] if the format is incorrect or the URI is unknown.
  factory NamespacePrefixedTag.fromClarkName(String clarkName) {
    if (!clarkName.startsWith('{') || !clarkName.contains('}')) {
      throw ArgumentError(
          "Clark name must be in the form '{uri}localPart', got '$clarkName'");
    }
    final uriEndIndex = clarkName.indexOf('}');
    final uri = clarkName.substring(1, uriEndIndex);
    final localName = clarkName.substring(uriEndIndex + 1);

    final prefix = pfxmap[uri];
    if (prefix == null) {
      throw ArgumentError("Namespace URI '$uri' not found in pfxmap");
    }
    // Call the default constructor with the reconstructed nstag
    return NamespacePrefixedTag('$prefix:$localName');
  }

  /// The Clark notation representation (e.g., "{http://.../main}p").
  String get clarkName => '{$nsuri}$localPart';

  /// A single-member map containing the prefix and URI for this tag.
  /// Useful for passing to XML processing functions that require namespace maps.
  // Made this static as it doesn't depend on instance state once calculated
  Map<String, String> get namespaceMap => {nspfx: nsuri};

  /// Helper to extract prefix, includes validation. MUST BE STATIC.
  static String _extractPrefix(String nstag) {
    final parts = nstag.split(':');
    if (parts.length != 2 || parts[0].isEmpty || parts[1].isEmpty) {
      throw ArgumentError(
          "Tag name '$nstag' must be in the form 'prefix:localName'");
    }
    return parts[0];
  }

  /// Helper to extract local part. MUST BE STATIC.
  static String _extractLocalPart(String nstag) {
     final parts = nstag.split(':');
     // Assume validation happened in _extractPrefix if called correctly
     if (parts.length != 2) return ''; // Or handle error differently
    return parts[1];
  }

  /// Helper to lookup URI, includes validation. MUST BE STATIC.
  static String _lookupUri(String prefix) {
    final uri = nsmap[prefix];
    if (uri == null) {
      throw ArgumentError("Namespace prefix '$prefix' not found in nsmap");
    }
    return uri;
  }

  @override
  String toString() => nstag;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is NamespacePrefixedTag &&
          runtimeType == other.runtimeType &&
          nstag == other.nstag;

  @override
  int get hashCode => nstag.hashCode;
}


/// Returns a string containing namespace declarations (e.g., `xmlns:pfx="uri"`)
/// for each prefix in the [prefixes] list.
///
/// Handy for adding required namespace declarations to a root element.
/// Throws [ArgumentError] if any prefix is not found in [nsmap].
String nsdecls(List<String> prefixes) {
  return prefixes.map((pfx) {
    final uri = nsmap[pfx];
    if (uri == null) {
      throw ArgumentError("Namespace prefix '$pfx' not found in nsmap");
    }
    return 'xmlns:$pfx="$uri"';
  }).join(' ');
}

/// Returns a subset of the main [nsmap] containing only the mappings
/// specified by the provided [nspfxs].
Map<String, String> nspfxmap(List<String> nspfxs) {
  final map = <String, String>{};
  for (final pfx in nspfxs) {
    final uri = nsmap[pfx];
    if (uri == null) {
      throw ArgumentError("Namespace prefix '$pfx' not found in nsmap");
    }
    map[pfx] = uri;
  }
  return map;
}

/// Converts a namespace-prefixed tag name (e.g., "w:p") into its
/// Clark notation equivalent ("{uri}localPart").
///
/// Throws [ArgumentError] if the tag format is incorrect or the prefix is unknown.
String qn(String tagName) {
  try {
    // Create an instance to access the calculated Clark name
    return NamespacePrefixedTag(tagName).clarkName;
  } on ArgumentError catch (e) {
    // Re-throw with potentially slightly more context if desired,
    // or just let the original argument error propagate.
    throw ArgumentError(
        "Could not get qualified name for '$tagName': ${e.message}");
  }
}
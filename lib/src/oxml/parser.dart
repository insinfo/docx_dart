/// Path: lib/src/oxml/parser.dart
/// Based on python-docx: docx/oxml/parser.py
///
/// XML parsing utilities for Open XML elements.

import 'package:xml/xml.dart';
// Note: BaseOxmlElement is not directly used here for parsing,
// but might be relevant for type hints or future integration.
// import 'xmlchemy.dart';
import 'ns.dart'; // For nsmap and qn function

// --- Differences from lxml/python-docx ---
// 1. No Direct Custom Class Parsing: package:xml parses into standard XmlElement
//    instances. Wrapping these elements into custom Dart classes (like BaseOxmlElement
//    subclasses) typically happens *after* parsing, based on the element tag.
// 2. No `register_element_cls`: Due to #1, there's no mechanism equivalent
//    to lxml's element class lookup registration during parsing.
// 3. Element Creation: XmlBuilder is used for robust element creation, which
//    handles namespaces differently than lxml's makeelement.

/// Parses an XML string or byte list (`List<int>`) into a root [XmlElement].
///
/// Assumes the input represents a single root element, possibly with children,
/// typical for an XML part in an OPC package.
/// Assumes UTF-8 encoding if input is `List<int>`.
///
/// Throws [XmlException] if XML parsing fails.
/// Throws [ArgumentError] if input is not [String] or [List<int>], or if
/// `List<int>` cannot be decoded as UTF-8.
XmlElement parseXml(dynamic xmlInput) {
  String xmlString;
  if (xmlInput is String) {
    xmlString = xmlInput;
  } else if (xmlInput is List<int>) {
    // Assume UTF-8 encoding, which is standard for OPC XML parts.
    try {
      // Note: If BOM is present, dart:convert's utf8.decode might be safer,
      // but fromCharCodes often works for typical OOXML parts.
      xmlString = String.fromCharCodes(xmlInput);
      // Optionally remove BOM if present:
      if (xmlString.startsWith('\uFEFF')) {
        xmlString = xmlString.substring(1);
      }
    } catch (e) {
      throw ArgumentError("Input List<int> could not be decoded as UTF-8: $e");
    }
  } else {
    throw ArgumentError(
        "Input must be String or List<int> (UTF-8 bytes), got ${xmlInput.runtimeType}");
  }

  try {
    // Parse the string into a document object.
    final document = XmlDocument.parse(xmlString);

    // Find and return the root element.
    // Throws StateError if no root element exists (shouldn't happen for valid XML parts).
    return document.rootElement;
  } on XmlException catch (e) {
    // Rethrow parsing exceptions for the caller to handle.
    print("Error parsing XML: $e"); // Optional: log the error
    rethrow;
  }
}

/// Creates a 'loose' [XmlElement] (not attached to a document tree)
/// with the specified prefixed tag name [nsptagStr].
///
/// This simulates lxml's `makeelement` used in `python-docx`'s `OxmlElement`.
///
/// Example:
/// ```dart
/// final p = OxmlElement("w:p",
///     attrs: {"rsidR": "123"},
///     nsdecls: {"w": nsmap['w']!} // Or let it be inferred
/// );
/// ```
///
/// - [nsptagStr] must be in the format "prefix:localName" (e.g., "w:p").
/// - [attrs] is an optional map of attribute names (can be prefixed like "r:id") to string values.
/// - [nsdecls] is an optional map of namespace prefix declarations (e.g., {"w": "http://..."}).
///   If [nsdecls] is not provided, a declaration for the prefix in [nsptagStr] is added automatically.
///
/// Throws [ArgumentError] if [nsptagStr] format is wrong, or if any prefix
/// (in the tag or attributes) is not found in the predefined `nsmap`.
XmlElement OxmlElement(String nsptagStr,
    {Map<String, String>? attrs, Map<String, String>? nsdecls}) {
  // 1. Validate and parse the main tag string
  final NamespacePrefixedTag tag;
  try {
    tag = nsptagStr.startsWith('{')
        ? NamespacePrefixedTag.fromClarkName(nsptagStr)
        : NamespacePrefixedTag(nsptagStr);
  } on ArgumentError catch (e) {
    throw ArgumentError(
        "Invalid nsptagStr '$nsptagStr' for OxmlElement: ${e.message}");
  }

  // 2. Prepare attributes with XmlName keys
  final processedAttrs = <XmlName, String>{};
  if (attrs != null) {
    attrs.forEach((attrName, attrValue) {
      final attrParts = attrName.split(':');
      String attrPrefix;
      String attrLocalName;
      String? attrUri;
      if (attrName.startsWith('{') && attrName.contains('}')) {
        final tag =
            NamespacePrefixedTag.fromClarkName(attrName);
        attrPrefix = tag.nspfx;
        attrLocalName = tag.localPart;
        attrUri = tag.nsuri;
      } else if (attrParts.length == 2) {
        // Prefixed attribute (e.g., "r:id", "xml:space")
        attrPrefix = attrParts[0];
        attrLocalName = attrParts[1];
        attrUri = nsmap[attrPrefix]; // Lookup URI from prefix
        if (attrUri == null) {
          // Only allow 'xml' prefix without explicit nsmap entry
          if (attrPrefix == 'xml') {
            attrUri = nsmap['xml']; // Use standard XML namespace URI
          } else {
            throw ArgumentError(
                "Namespace prefix '$attrPrefix' for attribute '$attrName' not found in nsmap");
          }
        }
      } else {
        // Attribute without prefix - belongs to no namespace
        attrPrefix = '';
        attrLocalName = attrName;
        attrUri = null;
      }
      processedAttrs[XmlName(attrLocalName, attrPrefix)] =
          attrValue; // Use XmlName(local, prefix)
    });
  }

  // 3. Determine namespace declarations to add to the element itself
  //    XmlBuilder needs URI -> Prefix mapping
  final effectiveNsDecls = nsdecls ?? {tag.nspfx: tag.nsuri};
  final namespaces = <String, String>{}; // Builder expects {URI: Prefix}
  effectiveNsDecls.forEach((pfx, nsUri) {
    namespaces[nsUri] = pfx;
  });

  // 4. Use XmlBuilder for robust element creation
  final builder = XmlBuilder();

  // Build the element with its namespace, attributes, and declarations
  // Note: XmlBuilder API can vary slightly. This works for recent versions.
  builder.element(
    tag.localPart, // The local name (e.g., 'p')
    namespace: tag.nsuri, // The element's namespace URI
    namespaces: namespaces, // Declarations {URI: Prefix} to add
    attributes: processedAttrs.map(
        (name, value) => MapEntry(name.qualified, value)), // {QName: Value}
  );

  // 5. Build the XML node structure
  final node = builder.buildFragment();

  // 6. Extract and return the created element
  if (node.children.isEmpty || node.children.first is! XmlElement) {
    // This should not happen if the builder worked correctly
    throw StateError(
        'Internal error: XmlBuilder did not create an element for $nsptagStr');
  }

  final element = node.children.first as XmlElement;
  element.parent?.children.remove(element);
  return element;
}


/// Creates a 'loose' [XmlElement] (not attached to a document tree)
/// with the specified prefixed tag name [nsptagStr], attributes, namespace
/// declarations, AND optional children.
///
/// This is a variant of [OxmlElement] specifically for cases where initial
/// children need to be added during creation (like in static `create` methods).
///
/// Example:
/// ```dart
/// final p = OxmlElementWithChildren("w:tc", children: [ CT_P.create() ]);
/// ```
///
/// - [nsptagStr], [attrs], [nsdecls]: Same as [OxmlElement].
/// - [children]: An optional list of [XmlElement]s to add as children.
///
/// Throws errors similar to [OxmlElement] for invalid inputs.
XmlElement OxmlElementWithChildren(String nsptagStr,
    {Map<String, String>? attrs,
    Map<String, String>? nsdecls,
    List<XmlElement>? children}) { // Added children parameter
  // 1. Create the parent element using the original function
  //    This handles tag parsing, attributes, and namespace declarations correctly.
  final parentElement = OxmlElement(nsptagStr, attrs: attrs, nsdecls: nsdecls);

  // 2. Add provided children, if any
  if (children != null && children.isNotEmpty) {
    for (final child in children) {
      // IMPORTANT: Detach child from previous parent if necessary
      // package:xml throws if adding a node that already has a parent.
      child.parent?.children.remove(child);
      parentElement.children.add(child);
    }
  }

  // 3. Return the parent element with its children added
  return parentElement;
}


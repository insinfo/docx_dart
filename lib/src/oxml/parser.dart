/// Path: lib/src/oxml/parser.dart
/// Based on python-docx: docx/oxml/parser.py
///
/// XML parsing utilities for Open XML elements.



import 'package:xml/xml.dart';
// Note: BaseOxmlElement is not directly used here for parsing,
// but might be relevant for type hints or future integration.
// import 'xmlchemy.dart';
import 'ns.dart'; // For nsmap and qn function

// The lxml custom parser with class lookup (`element_class_lookup`) and the
// `register_element_cls` function do not have direct equivalents in `package:xml`.
// In Dart, parsing typically yields standard `XmlElement` instances.
// Wrapping these elements in specific `BaseOxmlElement` subclasses happens *after*
// parsing, usually based on the element's tag name, often within factory
// constructors or specific methods of parent elements.

/// Parses an XML string or byte list into a root [XmlElement].
///
/// Throws [XmlException] if parsing fails.
/// Throws [ArgumentError] if input is not [String] or [List<int>].
XmlElement parseXml(dynamic xmlInput) {
  String xmlString;
  if (xmlInput is String) {
    xmlString = xmlInput;
  } else if (xmlInput is List<int>) {
    // Assume UTF-8 encoding, which is standard for OPC XML parts.
    try {
      xmlString = String.fromCharCodes(xmlInput);
    } catch (e) {
      throw ArgumentError("Input List<int> could not be decoded as UTF-8: $e");
    }
  } else {
    throw ArgumentError(
        "Input must be String or List<int> (UTF-8 bytes), got ${xmlInput.runtimeType}");
  }

  try {
    // ParseMode.element prevents parsing as a full document if only an element is provided
    // trimWhitespace: true removes ignorable whitespace text nodes
    // entityMapping: XmlDefaultEntityMapping.xml() prevents entity expansion for security
    final document = XmlDocument.parse(
      xmlString,
      // entityMapping: XmlDefaultEntityMapping.xml(), // Recommended for security
      // trimWhitespace: true, // Can be helpful, similar to remove_blank_text
    );
    return document.rootElement;
  } on XmlException catch (e) {
    print("Error parsing XML: $e");
    rethrow; // Propagate the parsing error
  }
}

/// Creates a 'loose' [XmlElement] (not attached to a document tree)
/// with the specified prefixed tag name [nsptagStr].
///
/// Example: `OxmlElement("w:p", attrs: {"rsidR": "123"}, nsdecls: {"w": nsmap['w']!})`
///
/// [nsptagStr] must be in the format "prefix:localName" (e.g., "w:p").
/// [attrs] is an optional map of attribute names (can be prefixed like "r:id") to values.
/// [nsdecls] is an optional map of namespace prefix declarations (e.g., {"w": "http://..."}).
/// If [nsdecls] is not provided, a declaration for the prefix in [nsptagStr] is added.
/// Throws [ArgumentError] if the prefix in [nsptagStr] or attribute names is not found in `nsmap`.
XmlElement OxmlElement(String nsptagStr,
    {Map<String, String>? attrs, Map<String, String>? nsdecls}) {
  final parts = nsptagStr.split(':');
  if (parts.length != 2 || parts[0].isEmpty || parts[1].isEmpty) {
    throw ArgumentError(
        "nsptagStr '$nsptagStr' must be in the form 'prefix:localName'");
  }
  final prefix = parts[0];
  final localName = parts[1];
  final uri = nsmap[prefix];
  if (uri == null) {
    throw ArgumentError(
        "Namespace prefix '$prefix' from tag '$nsptagStr' not found in nsmap");
  }

  // Prepare attributes with proper namespaces
  final processedAttrs = <XmlName, String>{};
  if (attrs != null) {
    attrs.forEach((attrName, attrValue) {
      final attrParts = attrName.split(':');
      String attrPrefix;
      String attrLocalName;
      String? attrUri;
      if (attrParts.length == 2) {
        attrPrefix = attrParts[0];
        attrLocalName = attrParts[1];
        attrUri = nsmap[attrPrefix];
        if (attrUri == null && attrPrefix != 'xml') {
          // Allow xml prefix implicitly
          throw ArgumentError(
              "Namespace prefix '$attrPrefix' for attribute '$attrName' not found in nsmap");
        }
        // Handle xml: prefix specifically if needed (often implicit)
        if (attrPrefix == 'xml') attrUri = nsmap['xml'];

      } else {
        // Attribute without prefix - no namespace
        attrLocalName = attrName;
        attrUri = null;
      }
       processedAttrs[XmlName(attrLocalName, attrUri)] = attrValue;
    });
  }


  // Determine namespace declarations to add
  final effectiveNsDecls = nsdecls ?? {prefix: uri};
  final namespaces = <String, String>{};
   effectiveNsDecls.forEach((pfx, nsUri) {
      namespaces[nsUri] = pfx; // package:xml uses uri -> prefix for builder
   });


  // Use XmlBuilder for robust element creation
  final builder = XmlBuilder();

  // Note: The builder API changed slightly over versions.
  // This assumes a version where attributes and namespaces are passed to `element`.
  // If using an older version, you might need builder.attribute(...) calls after builder.element(...).
  builder.element(
    localName,
    namespace: uri,
    namespaces: namespaces, // Pass URI->Prefix map
    attributes: processedAttrs.map((name, value) => MapEntry(name.qualified, value)), // Pass QName->Value map
  );

  // Build the node and return the element
  final node = builder.buildFragment();

  // Check if the fragment contains an element (it should)
  if (node.children.isEmpty || node.children.first is! XmlElement) {
    throw StateError('XmlBuilder did not create an element for $nsptagStr');
  }

  return node.children.first as XmlElement;
}
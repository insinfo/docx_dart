/// Path: lib/src/oxml/coreprops.dart
/// Based on python-docx: docx/oxml/coreprops.py
/// Custom element class for the <cp:coreProperties> element.

import 'package:xml/xml.dart';
import 'ns.dart';
import 'parser.dart';
import 'xmlchemy.dart';

/// `<cp:coreProperties>` element, the root element of the Core Properties part.
/// Stored as `/docProps/core.xml`. Implements many of the Dublin Core document metadata elements.
class CT_CoreProperties extends BaseOxmlElement {
  CT_CoreProperties(super.element);

  // Define tag names for properties
  static const String _cpCategoryTag = "cp:category";
  static const String _cpContentStatusTag = "cp:contentStatus";
  static const String _dctermsCreatedTag = "dcterms:created";
  static const String _dcCreatorTag = "dc:creator";
  static const String _dcDescriptionTag = "dc:description";
  static const String _dcIdentifierTag = "dc:identifier";
  static const String _cpKeywordsTag = "cp:keywords";
  static const String _dcLanguageTag = "dc:language";
  static const String _cpLastModifiedByTag = "cp:lastModifiedBy";
  static const String _cpLastPrintedTag = "cp:lastPrinted";
  static const String _dctermsModifiedTag = "dcterms:modified";
  static const String _cpRevisionTag = "cp:revision";
  static const String _dcSubjectTag = "dc:subject";
  static const String _dcTitleTag = "dc:title";
  static const String _cpVersionTag = "cp:version";

  // Define successor lists for insertion order (can be simplified/omitted if order is less strict)
  static const List<String> _successors = []; // Empty means append

  /// Returns a new `<cp:coreProperties>` element wrapped in this class.
  static CT_CoreProperties newCoreProperties() {
    final element = OxmlElement('cp:coreProperties', nsdecls: {
      // Passando as declarações de namespace necessárias
      'cp': nsmap['cp']!,
      'dc': nsmap['dc']!,
      'dcterms': nsmap['dcterms']!,
      'xsi': nsmap['xsi']!, // Adicionar xsi aqui também é uma boa prática
    });
    // Add required namespaces (nsdecls is not directly applicable here)
    // The OxmlElement factory should ideally handle adding namespace decls.
    // If not, they might need to be added manually or ensured by the part writer.
    return CT_CoreProperties(element);
  }

  /// The text in the `dc:creator` child element, or empty string if not present.
  String get authorText => _getTextOfElement(_dcCreatorTag);
  set authorText(String value) => _setElementText(_dcCreatorTag, value);

  /// The text in the `cp:category` child element, or empty string if not present.
  String get categoryText => _getTextOfElement(_cpCategoryTag);
  set categoryText(String value) => _setElementText(_cpCategoryTag, value);

  /// The text in the `dc:description` child element (used for comments), or empty string if not present.
  String get commentsText => _getTextOfElement(_dcDescriptionTag);
  set commentsText(String value) => _setElementText(_dcDescriptionTag, value);

  /// The text in the `cp:contentStatus` child element, or empty string if not present.
  String get contentStatusText => _getTextOfElement(_cpContentStatusTag);
  set contentStatusText(String value) =>
      _setElementText(_cpContentStatusTag, value);

  /// The [DateTime] value of the `dcterms:created` child element, or `null` if not present or invalid format.
  DateTime? get createdDatetime => _getDateTimeOfElement(_dctermsCreatedTag);
  set createdDatetime(DateTime? value) =>
      _setElementDateTime(_dctermsCreatedTag, value);

  /// The text in the `dc:identifier` child element, or empty string if not present.
  String get identifierText => _getTextOfElement(_dcIdentifierTag);
  set identifierText(String value) => _setElementText(_dcIdentifierTag, value);

  /// The text in the `cp:keywords` child element, or empty string if not present.
  String get keywordsText => _getTextOfElement(_cpKeywordsTag);
  set keywordsText(String value) => _setElementText(_cpKeywordsTag, value);

  /// The text in the `dc:language` child element, or empty string if not present.
  String get languageText => _getTextOfElement(_dcLanguageTag);
  set languageText(String value) => _setElementText(_dcLanguageTag, value);

  /// The text in the `cp:lastModifiedBy` child element, or empty string if not present.
  String get lastModifiedByText => _getTextOfElement(_cpLastModifiedByTag);
  set lastModifiedByText(String value) =>
      _setElementText(_cpLastModifiedByTag, value);

  /// The [DateTime] value of the `cp:lastPrinted` child element, or `null` if not present or invalid format.
  DateTime? get lastPrintedDatetime => _getDateTimeOfElement(_cpLastPrintedTag);
  set lastPrintedDatetime(DateTime? value) =>
      _setElementDateTime(_cpLastPrintedTag, value);

  /// The [DateTime] value of the `dcterms:modified` child element, or `null` if not present or invalid format.
  DateTime? get modifiedDatetime => _getDateTimeOfElement(_dctermsModifiedTag);
  set modifiedDatetime(DateTime? value) =>
      _setElementDateTime(_dctermsModifiedTag, value);

  /// The integer value of the `cp:revision` child element, or 0 if not present, non-integer, or negative.
  int get revisionNumber {
    final revisionElement = childOrNull(_cpRevisionTag);
    if (revisionElement == null || revisionElement.innerText.isEmpty) {
      return 0;
    }
    final revision = int.tryParse(revisionElement.innerText) ?? 0;
    return (revision < 0) ? 0 : revision;
  }

  set revisionNumber(int value) {
    if (value < 1) {
      throw ArgumentError(
          "revision property requires a positive int, got '$value'");
    }
    final revisionElement =
        getOrAddChild(_cpRevisionTag, _successors, _revisionFactory);
    revisionElement.innerText = value.toString();
  }

  /// The text in the `dc:subject` child element, or empty string if not present.
  String get subjectText => _getTextOfElement(_dcSubjectTag);
  set subjectText(String value) => _setElementText(_dcSubjectTag, value);

  /// The text in the `dc:title` child element, or empty string if not present.
  String get titleText => _getTextOfElement(_dcTitleTag);
  set titleText(String value) => _setElementText(_dcTitleTag, value);

  /// The text in the `cp:version` child element, or empty string if not present.
  String get versionText => _getTextOfElement(_cpVersionTag);
  set versionText(String value) => _setElementText(_cpVersionTag, value);

  // --- Private Helper Methods ---

  XmlElement _revisionFactory() => OxmlElement(_cpRevisionTag);

  /// Gets the text content of a child element by tag name, returns "" if absent.
  String _getTextOfElement(String tagName) {
    final child = childOrNull(tagName);
    return child?.innerText ?? "";
  }

  /// Sets the text content of a child element by tag name. Adds element if needed.
  /// Limits text length to 255 characters.
  void _setElementText(String tagName, String? value) {
    if (value == null || value.isEmpty) {
      removeChild(tagName);
      return;
    }
    if (value.length > 255) {
      throw ArgumentError(
          "exceeded 255 char limit for property '$tagName', got ${value.length} chars");
    }
    final child =
        getOrAddChild(tagName, _successors, () => OxmlElement(tagName));
    child.innerText = value;
  }

  /// Parses the text of a child element as a W3CDTF DateTime (ISO 8601).
  /// Returns null if element is absent or text is not a valid format.
  DateTime? _getDateTimeOfElement(String tagName) {
    final child = childOrNull(tagName);
    final text = child?.innerText;
    if (text == null || text.isEmpty) {
      return null;
    }
    try {
      // Use `DateTime.parse` which handles ISO 8601 including 'Z' for UTC
      // It might not handle *all* W3CDTF variants (like year only), but covers common cases.
      final dt = DateTime.parse(text);
      // Ensure it's treated as UTC if 'Z' was present or no offset specified
      return dt.toUtc();
    } catch (e) {
      print("Warn: Could not parse W3CDTF datetime string '$text': $e");
      return null; // Ignore invalid datetime strings
    }
  }

  /// Sets the text of a child element to a W3CDTF DateTime string (ISO 8601 UTC format).
  /// Adds element if needed. Adds xsi:type attribute for created/modified.
  void _setElementDateTime(String tagName, DateTime? value) {
    if (value == null) {
      removeChild(tagName);
      return;
    }

    // Format as ISO 8601 UTC with 'Z'
    // Example: 2023-10-27T10:30:00.123Z
    final dtStr = value.toUtc().toIso8601String();

    final child =
        getOrAddChild(tagName, _successors, () => OxmlElement(tagName));
    child.innerText = dtStr;

    // Add xsi:type attribute for dcterms:created and dcterms:modified
    if (tagName == _dctermsCreatedTag || tagName == _dctermsModifiedTag) {
      // Ensure the xsi namespace is declared on the root or this element.
      // This might require handling at a higher level (e.g., when saving the part)
      // or adding it directly here if the OxmlElement factory doesn't handle it.
      // For simplicity here, we just set the attribute.
      child.setAttribute('type', 'dcterms:W3CDTF', namespace: nsmap['xsi']);
      // Ensure xsi prefix is defined in the nsmap used by the element/document.
    } else {
      // Remove xsi:type if it exists and is not needed
      child.removeAttribute('type', namespace: nsmap['xsi']);
    }
  }
}

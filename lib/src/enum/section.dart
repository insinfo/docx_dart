// docx/enum/section.dart

// Corresponds to Python file docx/enum/section.py

/// Specifies one of the three possible header/footer definitions for a section.
///
/// Alias: `WD_HEADER_FOOTER`
///
/// For internal use only; not part of the `python-docx` API.
///
/// MS API name: `WdHeaderFooterIndex`
/// URL: https://docs.microsoft.com/en-us/office/vba/api/word.wdheaderfooterindex
enum WD_HEADER_FOOTER_INDEX {
  /// Header for odd pages or all if no even header.
  PRIMARY(1, "default", "Header for odd pages or all if no even header."),

  /// Header for first page of section.
  FIRST_PAGE(2, "first", "Header for first page of section."),

  /// Header for even pages of recto/verso section.
  EVEN_PAGE(3, "even", "Header for even pages of recto/verso section.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_HEADER_FOOTER_INDEX(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_HEADER_FOOTER_INDEX fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_HEADER_FOOTER_INDEX.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_HEADER_FOOTER_INDEX has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_HEADER_FOOTER_INDEX? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_HEADER_FOOTER_INDEX].
typedef WD_HEADER_FOOTER = WD_HEADER_FOOTER_INDEX;


/// Specifies the page layout orientation.
///
/// Alias: `WD_ORIENT`
///
/// Example:
/// ```dart
/// // Assuming 'section' is an object with an 'orientation' property
/// // section.orientation = WD_ORIENTATION.LANDSCAPE;
/// ```
///
/// MS API name: `WdOrientation`
/// MS API URL: http://msdn.microsoft.com/en-us/library/office/ff837902.aspx
enum WD_ORIENTATION {
  /// Portrait orientation.
  PORTRAIT(0, "portrait", "Portrait orientation."),

  /// Landscape orientation.
  LANDSCAPE(1, "landscape", "Landscape orientation.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_ORIENTATION(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_ORIENTATION fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_ORIENTATION.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_ORIENTATION has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_ORIENTATION? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_ORIENTATION].
typedef WD_ORIENT = WD_ORIENTATION;


/// Specifies the start type of a section break.
///
/// Alias: `WD_SECTION`
///
/// Example:
/// ```dart
/// // Assuming 'section' is an object with a 'startType' property
/// // section.startType = WD_SECTION_START.NEW_PAGE;
/// ```
///
/// MS API name: `WdSectionStart`
/// MS API URL: http://msdn.microsoft.com/en-us/library/office/ff840975.aspx
enum WD_SECTION_START {
  /// Continuous section break.
  CONTINUOUS(0, "continuous", "Continuous section break."),

  /// New column section break.
  NEW_COLUMN(1, "nextColumn", "New column section break."),

  /// New page section break.
  NEW_PAGE(2, "nextPage", "New page section break."),

  /// Even pages section break.
  EVEN_PAGE(3, "evenPage", "Even pages section break."),

  /// Section begins on next odd page.
  ODD_PAGE(4, "oddPage", "Section begins on next odd page.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_SECTION_START(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_SECTION_START fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_SECTION_START.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_SECTION_START has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_SECTION_START? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_SECTION_START].
typedef WD_SECTION = WD_SECTION_START;
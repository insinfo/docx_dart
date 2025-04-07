// docx/enum/table.dart

// Corresponds to Python file docx/enum/table.py

/// Specifies the vertical alignment of text in one or more cells of a table.
///
/// Alias: `WD_ALIGN_VERTICAL`
///
/// Example:
/// ```dart
/// // Assuming 'cell' is an object with a 'verticalAlignment' property
/// // cell.verticalAlignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM;
/// ```
///
/// MS API name: `WdCellVerticalAlignment`
///
/// https://msdn.microsoft.com/en-us/library/office/ff193345.aspx
enum WD_CELL_VERTICAL_ALIGNMENT {
  /// Text is aligned to the top border of the cell.
  TOP(0, "top", "Text is aligned to the top border of the cell."),

  /// Text is aligned to the center of the cell.
  CENTER(1, "center", "Text is aligned to the center of the cell."),

  /// Text is aligned to the bottom border of the cell.
  BOTTOM(3, "bottom", "Text is aligned to the bottom border of the cell."),

  /// This is an option in the OpenXml spec, but not in Word itself. It's not
  /// clear what Word behavior this setting produces. If you find out please
  /// let us know and we'll update this documentation. Otherwise, probably best
  /// to avoid this option.
  BOTH(
    101,
    "both",
    "This is an option in the OpenXml spec, but not in Word itself. It's not "
        "clear what Word behavior this setting produces. If you find out please "
        "let us know and we'll update this documentation. Otherwise, probably best "
        "to avoid this option.",
  );

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_CELL_VERTICAL_ALIGNMENT(
      this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_CELL_VERTICAL_ALIGNMENT fromXml(String? xmlValue) {
    if (xmlValue == null) {
      throw ArgumentError(
          'WD_CELL_VERTICAL_ALIGNMENT.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError(
        'WD_CELL_VERTICAL_ALIGNMENT has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_CELL_VERTICAL_ALIGNMENT? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_CELL_VERTICAL_ALIGNMENT].
typedef WD_ALIGN_VERTICAL = WD_CELL_VERTICAL_ALIGNMENT;

/// Specifies the rule for determining the height of a table row.
///
/// Alias: `WD_ROW_HEIGHT`
///
/// Example:
/// ```dart
/// // Assuming 'row' is an object with a 'heightRule' property
/// // row.heightRule = WD_ROW_HEIGHT_RULE.EXACTLY;
/// ```
///
/// MS API name: `WdRowHeightRule`
///
/// https://msdn.microsoft.com/en-us/library/office/ff193620.aspx
enum WD_ROW_HEIGHT_RULE {
  /// The row height is adjusted to accommodate the tallest value in the row.
  AUTO(
    0,
    "auto",
    "The row height is adjusted to accommodate the tallest value in the row.",
  ),

  /// The row height is at least a minimum specified value.
  AT_LEAST(
      1, "atLeast", "The row height is at least a minimum specified value."),

  /// The row height is an exact value.
  EXACTLY(2, "exact", "The row height is an exact value.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_ROW_HEIGHT_RULE(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_ROW_HEIGHT_RULE fromXml(String? xmlValue) {
    if (xmlValue == null) {
      throw ArgumentError(
          'WD_ROW_HEIGHT_RULE.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError(
        'WD_ROW_HEIGHT_RULE has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_ROW_HEIGHT_RULE? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_ROW_HEIGHT_RULE].
typedef WD_ROW_HEIGHT = WD_ROW_HEIGHT_RULE;

/// Specifies table justification type.
///
/// Example:
/// ```dart
/// // Assuming 'table' is an object with an 'alignment' property
/// // table.alignment = WD_TABLE_ALIGNMENT.CENTER;
/// ```
///
/// MS API name: `WdRowAlignment`
///
/// http://office.microsoft.com/en-us/word-help/HV080607259.aspx
enum WD_TABLE_ALIGNMENT {
  /// Left-aligned
  LEFT(0, "left", "Left-aligned"),

  /// Center-aligned.
  CENTER(1, "center", "Center-aligned."),

  /// Right-aligned.
  RIGHT(2, "right", "Right-aligned.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_TABLE_ALIGNMENT(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_TABLE_ALIGNMENT fromXml(String? xmlValue) {
    if (xmlValue == null) {
      throw ArgumentError(
          'WD_TABLE_ALIGNMENT.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError(
        'WD_TABLE_ALIGNMENT has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_TABLE_ALIGNMENT? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Specifies the direction in which an application orders cells in the
/// specified table or row.
///
/// Example:
/// ```dart
/// // Assuming 'table' is an object with a 'direction' property
/// // table.direction = WD_TABLE_DIRECTION.RTL;
/// ```
///
/// MS API name: `WdTableDirection`
///
/// http://msdn.microsoft.com/en-us/library/ff835141.aspx
enum WD_TABLE_DIRECTION {
  /// The table or row is arranged with the first column in the leftmost position.
  LTR(
    0,
    "The table or row is arranged with the first column in the leftmost position.",
  ),

  /// The table or row is arranged with the first column in the rightmost position.
  RTL(
    1,
    "The table or row is arranged with the first column in the rightmost position.",
  );

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_TABLE_DIRECTION(this.intValue, this.description);

  @override
  String toString() => '$name ($intValue)';
}

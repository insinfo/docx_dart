// docx/enum/text.dart

// Corresponds to Python file docx/enum/text.py

/// Specifies paragraph justification type.
///
/// Alias: `WD_ALIGN_PARAGRAPH`
///
/// Example:
/// ```dart
/// // Assuming 'paragraph' is an object with an 'alignment' property
/// // paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER;
/// ```
enum WD_PARAGRAPH_ALIGNMENT {
  /// Left-aligned
  LEFT(0, "left", "Left-aligned"),
  /// Center-aligned.
  CENTER(1, "center", "Center-aligned."),
  /// Right-aligned.
  RIGHT(2, "right", "Right-aligned."),
  /// Fully justified.
  JUSTIFY(3, "both", "Fully justified."),
  /// Paragraph characters are distributed to fill entire width of paragraph.
  DISTRIBUTE(
    4,
    "distribute",
    "Paragraph characters are distributed to fill entire width of paragraph.",
  ),
  /// Justified with a medium character compression ratio.
  JUSTIFY_MED(
    5,
    "mediumKashida",
    "Justified with a medium character compression ratio.",
  ),
  /// Justified with a high character compression ratio.
  JUSTIFY_HI(
    7,
    "highKashida",
    "Justified with a high character compression ratio.",
  ),
  /// Justified with a low character compression ratio.
  JUSTIFY_LOW(8, "lowKashida", "Justified with a low character compression ratio."),
  /// Justified according to Thai formatting layout.
  THAI_JUSTIFY(
    9,
    "thaiDistribute",
    "Justified according to Thai formatting layout.",
  );

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_PARAGRAPH_ALIGNMENT(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_PARAGRAPH_ALIGNMENT fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_PARAGRAPH_ALIGNMENT.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_PARAGRAPH_ALIGNMENT has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_PARAGRAPH_ALIGNMENT? value) {
    return value?.xmlValue;
  }

   @override
   String toString() => '$name ($intValue)';
}

/// Alias for [WD_PARAGRAPH_ALIGNMENT].
typedef WD_ALIGN_PARAGRAPH = WD_PARAGRAPH_ALIGNMENT;


/// Corresponds to WdBreakType enumeration.
///
/// Alias: `WD_BREAK`
///
/// http://msdn.microsoft.com/en-us/library/office/ff195905.aspx
enum WD_BREAK_TYPE {
  COLUMN(8, 'Column break'),
  LINE(6, 'Line break'),
  LINE_CLEAR_LEFT(9, 'Line break clearing left'),
  LINE_CLEAR_RIGHT(10, 'Line break clearing right'),
  LINE_CLEAR_ALL(11, 'Line break clearing all (added for consistency, not in MS version)'),
  PAGE(7, 'Page break'),
  SECTION_CONTINUOUS(3, 'Continuous section break'),
  SECTION_EVEN_PAGE(4, 'Even page section break'),
  SECTION_NEXT_PAGE(2, 'Next page section break'),
  SECTION_ODD_PAGE(5, 'Odd page section break'),
  TEXT_WRAPPING(11, 'Text wrapping break'); // Note: Value is same as LINE_CLEAR_ALL

  /// The integer value corresponding to the MS API enum member.
  final int intValue;
  /// The documentation string for the enum member.
  final String description;

  const WD_BREAK_TYPE(this.intValue, this.description);

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_BREAK_TYPE].
typedef WD_BREAK = WD_BREAK_TYPE;


/// Specifies a standard preset color to apply.
///
/// Alias: `WD_COLOR`
///
/// Used for font highlighting and perhaps other applications.
///
/// MS API name: `WdColorIndex`
/// URL: https://msdn.microsoft.com/EN-US/library/office/ff195343.aspx
enum WD_COLOR_INDEX {
  /// Color is inherited from the style hierarchy.
  INHERITED(-1, null, "Color is inherited from the style hierarchy."),
  /// Automatic color. Default; usually black.
  AUTO(0, "default", "Automatic color. Default; usually black."),
  /// Black color.
  BLACK(1, "black", "Black color."),
  /// Blue color
  BLUE(2, "blue", "Blue color"),
  /// Bright green color.
  BRIGHT_GREEN(4, "green", "Bright green color."),
  /// Dark blue color.
  DARK_BLUE(9, "darkBlue", "Dark blue color."),
  /// Dark red color.
  DARK_RED(13, "darkRed", "Dark red color."),
  /// Dark yellow color.
  DARK_YELLOW(14, "darkYellow", "Dark yellow color."),
  /// 25% shade of gray color.
  GRAY_25(16, "lightGray", "25% shade of gray color."),
  /// 50% shade of gray color.
  GRAY_50(15, "darkGray", "50% shade of gray color."),
  /// Green color.
  GREEN(11, "darkGreen", "Green color."),
  /// Pink color.
  PINK(5, "magenta", "Pink color."),
  /// Red color.
  RED(6, "red", "Red color."),
  /// Teal color.
  TEAL(10, "darkCyan", "Teal color."),
  /// Turquoise color.
  TURQUOISE(3, "cyan", "Turquoise color."),
  /// Violet color.
  VIOLET(12, "darkMagenta", "Violet color."),
  /// White color.
  WHITE(8, "white", "White color."),
  /// Yellow color.
  YELLOW(7, "yellow", "Yellow color.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member (nullable).
  final String? xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_COLOR_INDEX(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_COLOR_INDEX fromXml(String? xmlValue) {
     if (xmlValue == null) {
        // Handle null specifically for INHERITED
        return WD_COLOR_INDEX.INHERITED;
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_COLOR_INDEX has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_COLOR_INDEX? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_COLOR_INDEX].
typedef WD_COLOR = WD_COLOR_INDEX;


/// Specifies a line spacing format to be applied to a paragraph.
///
/// Example:
/// ```dart
/// // Assuming 'paragraph' is an object with a 'lineSpacingRule' property
/// // paragraph.lineSpacingRule = WD_LINE_SPACING.EXACTLY;
/// ```
///
/// MS API name: `WdLineSpacing`
///
/// URL: http://msdn.microsoft.com/en-us/library/office/ff844910.aspx
enum WD_LINE_SPACING {
  /// Single spaced (default).
  SINGLE(0, null, "Single spaced (default)."),
  /// Space-and-a-half line spacing.
  ONE_POINT_FIVE(1, null, "Space-and-a-half line spacing."),
  /// Double spaced.
  DOUBLE(2, null, "Double spaced."),
  /// Minimum line spacing is specified amount. Amount is specified separately.
  AT_LEAST(
    3,
    "atLeast",
    "Minimum line spacing is specified amount. Amount is specified separately.",
  ),
  /// Line spacing is exactly specified amount. Amount is specified separately.
  EXACTLY(
    4,
    "exact",
    "Line spacing is exactly specified amount. Amount is specified separately.",
  ),
  /// Line spacing is specified as multiple of line heights. Changing font size
  /// will change line spacing proportionately.
  MULTIPLE(
    5,
    "auto",
    "Line spacing is specified as multiple of line heights. Changing font size "
    "will change line spacing proportionately.",
  );

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member (nullable).
  final String? xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_LINE_SPACING(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  /// Handles the "UNMAPPED" values by throwing an error.
  static WD_LINE_SPACING fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_LINE_SPACING.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      // Skip SINGLE, ONE_POINT_FIVE, DOUBLE as they have no direct XML mapping here
      if (member.xmlValue != null && member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_LINE_SPACING has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_LINE_SPACING? value) {
    // Return null for members that don't have a direct XML mapping
    if (value == WD_LINE_SPACING.SINGLE ||
        value == WD_LINE_SPACING.ONE_POINT_FIVE ||
        value == WD_LINE_SPACING.DOUBLE) {
        return null;
    }
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}


/// Specifies the tab stop alignment to apply.
///
/// MS API name: `WdTabAlignment`
///
/// URL: https://msdn.microsoft.com/EN-US/library/office/ff195609.aspx
enum WD_TAB_ALIGNMENT {
  /// Left-aligned.
  LEFT(0, "left", "Left-aligned."),
  /// Center-aligned.
  CENTER(1, "center", "Center-aligned."),
  /// Right-aligned.
  RIGHT(2, "right", "Right-aligned."),
  /// Decimal-aligned.
  DECIMAL(3, "decimal", "Decimal-aligned."),
  /// Bar-aligned.
  BAR(4, "bar", "Bar-aligned."),
  /// List-aligned. (deprecated)
  LIST(6, "list", "List-aligned. (deprecated)"),
  /// Clear an inherited tab stop.
  CLEAR(101, "clear", "Clear an inherited tab stop."),
  /// Right-aligned. (deprecated)
  END(102, "end", "Right-aligned. (deprecated)"),
  /// Left-aligned. (deprecated)
  NUM(103, "num", "Left-aligned. (deprecated)"),
  /// Left-aligned. (deprecated)
  START(104, "start", "Left-aligned. (deprecated)");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_TAB_ALIGNMENT(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_TAB_ALIGNMENT fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_TAB_ALIGNMENT.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_TAB_ALIGNMENT has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_TAB_ALIGNMENT? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}


/// Specifies the character to use as the leader with formatted tabs.
///
/// MS API name: `WdTabLeader`
///
/// URL: https://msdn.microsoft.com/en-us/library/office/ff845050.aspx
enum WD_TAB_LEADER {
  /// Spaces. Default.
  SPACES(0, "none", "Spaces. Default."),
  /// Dots.
  DOTS(1, "dot", "Dots."),
  /// Dashes.
  DASHES(2, "hyphen", "Dashes."),
  /// Double lines.
  LINES(3, "underscore", "Double lines."),
  /// A heavy line.
  HEAVY(4, "heavy", "A heavy line."),
  /// A vertically-centered dot.
  MIDDLE_DOT(5, "middleDot", "A vertically-centered dot.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_TAB_LEADER(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_TAB_LEADER fromXml(String? xmlValue) {
     if (xmlValue == null) {
       throw ArgumentError('WD_TAB_LEADER.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_TAB_LEADER has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_TAB_LEADER? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}


/// Specifies the style of underline applied to a run of characters.
///
/// MS API name: `WdUnderline`
///
/// URL: http://msdn.microsoft.com/en-us/library/office/ff822388.aspx
enum WD_UNDERLINE {
  /// Inherit underline setting from containing paragraph.
  INHERITED(-1, null, "Inherit underline setting from containing paragraph."),

  /// No underline.
  ///
  /// This setting overrides any inherited underline value, so can be used to
  /// remove underline from a run that inherits underlining from its containing
  /// paragraph. Note this is not the same as assigning `null` to
  /// Run.underline. `null` is a valid assignment value, but causes the run to
  /// inherit its underline value. Assigning `WD_UNDERLINE.NONE` causes
  /// underlining to be unconditionally turned off.
  NONE(
    0,
    "none",
    "No underline.\n\nThis setting overrides any inherited underline value, so "
    "can be used to remove underline from a run that inherits underlining "
    "from its containing paragraph. Note this is not the same as assigning "
    "null to Run.underline. null is a valid assignment value, but causes "
    "the run to inherit its underline value. Assigning `WD_UNDERLINE.NONE` "
    "causes underlining to be unconditionally turned off.",
  ),

  /// A single line.
  ///
  /// Note that this setting is write-only in the sense that `true`
  /// (rather than `WD_UNDERLINE.SINGLE`) is returned for a run having this
  /// setting (in the original python-docx library). In Dart, you would check
  /// against `WD_UNDERLINE.SINGLE`.
  SINGLE(
    1,
    "single",
    "A single line.\n\nNote that this setting is write-only in the sense "
    "that true (rather than `WD_UNDERLINE.SINGLE`) is returned for a run "
    "having this setting.",
  ),

  /// Underline individual words only.
  WORDS(2, "words", "Underline individual words only."),
  /// A double line.
  DOUBLE(3, "double", "A double line."),
  /// Dots.
  DOTTED(4, "dotted", "Dots."),
  /// A single thick line.
  THICK(6, "thick", "A single thick line."),
  /// Dashes.
  DASH(7, "dash", "Dashes."),
  /// Alternating dots and dashes.
  DOT_DASH(9, "dotDash", "Alternating dots and dashes."),
  /// An alternating dot-dot-dash pattern.
  DOT_DOT_DASH(10, "dotDotDash", "An alternating dot-dot-dash pattern."),
  /// A single wavy line.
  WAVY(11, "wave", "A single wavy line."),
  /// Heavy dots.
  DOTTED_HEAVY(20, "dottedHeavy", "Heavy dots."),
  /// Heavy dashes.
  DASH_HEAVY(23, "dashedHeavy", "Heavy dashes."),
  /// Alternating heavy dots and heavy dashes.
  DOT_DASH_HEAVY(25, "dashDotHeavy", "Alternating heavy dots and heavy dashes."),
  /// An alternating heavy dot-dot-dash pattern.
  DOT_DOT_DASH_HEAVY(
    26,
    "dashDotDotHeavy",
    "An alternating heavy dot-dot-dash pattern.",
  ),
  /// A heavy wavy line.
  WAVY_HEAVY(27, "wavyHeavy", "A heavy wavy line."),
  /// Long dashes.
  DASH_LONG(39, "dashLong", "Long dashes."),
  /// A double wavy line.
  WAVY_DOUBLE(43, "wavyDouble", "A double wavy line."),
  /// Long heavy dashes.
  DASH_LONG_HEAVY(55, "dashLongHeavy", "Long heavy dashes.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member (nullable).
  final String? xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_UNDERLINE(this.intValue, this.xmlValue, this.description);

   /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_UNDERLINE fromXml(String? xmlValue) {
    if (xmlValue == null) {
      // Handle null specifically for INHERITED
      return WD_UNDERLINE.INHERITED;
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_UNDERLINE has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_UNDERLINE? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}
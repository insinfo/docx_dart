// docx/enum/dml.dart
// (Necessário para color.dart)

// Corresponds to Python file docx/enum/dml.py

/// Specifies the color specification scheme.
///
/// MS API name: `MsoColorType`
///
/// http://msdn.microsoft.com/en-us/library/office/ff864912(v=office.15).aspx
enum MSO_COLOR_TYPE {
  /// Color is specified by an |RGBColor| value.
  RGB(1, 'Color is specified by an |RGBColor| value.'),

  /// Color is one of the preset theme colors.
  THEME(2, 'Color is one of the preset theme colors.'),

  /// Color is determined automatically by the application.
  AUTO(101, 'Color is determined automatically by the application.');

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The documentation string for the enum member.
  final String description;

  const MSO_COLOR_TYPE(this.intValue, this.description);

  @override
  String toString() => '$name ($intValue)';
}

/// Indicates the Office theme color, one of those shown in the color gallery
/// on the formatting ribbon.
///
/// Alias: `MSO_THEME_COLOR`
///
/// Example:
/// ```dart
/// // Assuming font.color.themeColor is settable
/// // font.color.themeColor = MSO_THEME_COLOR_INDEX.ACCENT_1;
/// ```
///
/// MS API name: `MsoThemeColorIndex`
///
/// http://msdn.microsoft.com/en-us/library/office/ff860782(v=office.15).aspx
enum MSO_THEME_COLOR_INDEX {
  /// Indicates the color is not a theme color.
  NOT_THEME_COLOR(0, "UNMAPPED", "Indicates the color is not a theme color."),

  /// Specifies the Accent 1 theme color.
  ACCENT_1(5, "accent1", "Specifies the Accent 1 theme color."),

  /// Specifies the Accent 2 theme color.
  ACCENT_2(6, "accent2", "Specifies the Accent 2 theme color."),

  /// Specifies the Accent 3 theme color.
  ACCENT_3(7, "accent3", "Specifies the Accent 3 theme color."),

  /// Specifies the Accent 4 theme color.
  ACCENT_4(8, "accent4", "Specifies the Accent 4 theme color."),

  /// Specifies the Accent 5 theme color.
  ACCENT_5(9, "accent5", "Specifies the Accent 5 theme color."),

  /// Specifies the Accent 6 theme color.
  ACCENT_6(10, "accent6", "Specifies the Accent 6 theme color."),

  /// Specifies the Background 1 theme color.
  BACKGROUND_1(14, "background1", "Specifies the Background 1 theme color."),

  /// Specifies the Background 2 theme color.
  BACKGROUND_2(16, "background2", "Specifies the Background 2 theme color."),

  /// Specifies the Dark 1 theme color.
  DARK_1(1, "dark1", "Specifies the Dark 1 theme color."),

  /// Specifies the Dark 2 theme color.
  DARK_2(3, "dark2", "Specifies the Dark 2 theme color."),

  /// Specifies the theme color for a clicked hyperlink.
  FOLLOWED_HYPERLINK(
    12,
    "followedHyperlink",
    "Specifies the theme color for a clicked hyperlink.",
  ),

  /// Specifies the theme color for a hyperlink.
  HYPERLINK(11, "hyperlink", "Specifies the theme color for a hyperlink."),

  /// Specifies the Light 1 theme color.
  LIGHT_1(2, "light1", "Specifies the Light 1 theme color."),

  /// Specifies the Light 2 theme color.
  LIGHT_2(4, "light2", "Specifies the Light 2 theme color."),

  /// Specifies the Text 1 theme color.
  TEXT_1(13, "text1", "Specifies the Text 1 theme color."),

  /// Specifies the Text 2 theme color.
  TEXT_2(15, "text2", "Specifies the Text 2 theme color.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const MSO_THEME_COLOR_INDEX(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  ///
  /// Example:
  /// ```dart
  /// MSO_THEME_COLOR_INDEX? colorIndex = MSO_THEME_COLOR_INDEX.fromXml("accent1");
  /// print(colorIndex == MSO_THEME_COLOR_INDEX.ACCENT_1); // true
  /// ```
  /// Throws ArgumentError if no member matches `xmlValue`.
  static MSO_THEME_COLOR_INDEX fromXml(String? xmlValue) {
    if (xmlValue == null) {
       throw ArgumentError('MSO_THEME_COLOR_INDEX.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('MSO_THEME_COLOR_INDEX has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(MSO_THEME_COLOR_INDEX? value) {
    return value?.xmlValue;
  }

   @override
   String toString() => '$name ($intValue)';
}

/// Alias for [MSO_THEME_COLOR_INDEX].
typedef MSO_THEME_COLOR = MSO_THEME_COLOR_INDEX;
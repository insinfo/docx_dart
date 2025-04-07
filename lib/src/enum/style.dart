// docx/enum/style.dart

// Corresponds to Python file docx/enum/style.py

/// Specifies a built-in Microsoft Word style.
///
/// Alias: `WD_STYLE`
///
/// Example:
/// ```dart
/// // Assuming a Styles collection object 'styles'
/// // var style = styles[WD_BUILTIN_STYLE.BODY_TEXT];
/// ```
///
/// MS API name: `WdBuiltinStyle`
///
/// http://msdn.microsoft.com/en-us/library/office/ff835210.aspx
enum WD_BUILTIN_STYLE {
  /// Block Text.
  BLOCK_QUOTATION(-85, "Block Text."),

  /// Body Text.
  BODY_TEXT(-67, "Body Text."),

  /// Body Text 2.
  BODY_TEXT_2(-81, "Body Text 2."),

  /// Body Text 3.
  BODY_TEXT_3(-82, "Body Text 3."),

  /// Body Text First Indent.
  BODY_TEXT_FIRST_INDENT(-78, "Body Text First Indent."),

  /// Body Text First Indent 2.
  BODY_TEXT_FIRST_INDENT_2(-79, "Body Text First Indent 2."),

  /// Body Text Indent.
  BODY_TEXT_INDENT(-68, "Body Text Indent."),

  /// Body Text Indent 2.
  BODY_TEXT_INDENT_2(-83, "Body Text Indent 2."),

  /// Body Text Indent 3.
  BODY_TEXT_INDENT_3(-84, "Body Text Indent 3."),

  /// Book Title.
  BOOK_TITLE(-265, "Book Title."),

  /// Caption.
  CAPTION(-35, "Caption."),

  /// Closing.
  CLOSING(-64, "Closing."),

  /// Comment Reference.
  COMMENT_REFERENCE(-40, "Comment Reference."),

  /// Comment Text.
  COMMENT_TEXT(-31, "Comment Text."),

  /// Date.
  DATE(-77, "Date."),

  /// Default Paragraph Font.
  DEFAULT_PARAGRAPH_FONT(-66, "Default Paragraph Font."),

  /// Emphasis.
  EMPHASIS(-89, "Emphasis."),

  /// Endnote Reference.
  ENDNOTE_REFERENCE(-43, "Endnote Reference."),

  /// Endnote Text.
  ENDNOTE_TEXT(-44, "Endnote Text."),

  /// Envelope Address.
  ENVELOPE_ADDRESS(-37, "Envelope Address."),

  /// Envelope Return.
  ENVELOPE_RETURN(-38, "Envelope Return."),

  /// Footer.
  FOOTER(-33, "Footer."),

  /// Footnote Reference.
  FOOTNOTE_REFERENCE(-39, "Footnote Reference."),

  /// Footnote Text.
  FOOTNOTE_TEXT(-30, "Footnote Text."),

  /// Header.
  HEADER(-32, "Header."),

  /// Heading 1.
  HEADING_1(-2, "Heading 1."),

  /// Heading 2.
  HEADING_2(-3, "Heading 2."),

  /// Heading 3.
  HEADING_3(-4, "Heading 3."),

  /// Heading 4.
  HEADING_4(-5, "Heading 4."),

  /// Heading 5.
  HEADING_5(-6, "Heading 5."),

  /// Heading 6.
  HEADING_6(-7, "Heading 6."),

  /// Heading 7.
  HEADING_7(-8, "Heading 7."),

  /// Heading 8.
  HEADING_8(-9, "Heading 8."),

  /// Heading 9.
  HEADING_9(-10, "Heading 9."),

  /// HTML Acronym.
  HTML_ACRONYM(-96, "HTML Acronym."),

  /// HTML Address.
  HTML_ADDRESS(-97, "HTML Address."),

  /// HTML Cite.
  HTML_CITE(-98, "HTML Cite."),

  /// HTML Code.
  HTML_CODE(-99, "HTML Code."),

  /// HTML Definition.
  HTML_DFN(-100, "HTML Definition."),

  /// HTML Keyboard.
  HTML_KBD(-101, "HTML Keyboard."),

  /// Normal (Web).
  HTML_NORMAL(-95, "Normal (Web)."),

  /// HTML Preformatted.
  HTML_PRE(-102, "HTML Preformatted."),

  /// HTML Sample.
  HTML_SAMP(-103, "HTML Sample."),

  /// HTML Typewriter.
  HTML_TT(-104, "HTML Typewriter."),

  /// HTML Variable.
  HTML_VAR(-105, "HTML Variable."),

  /// Hyperlink.
  HYPERLINK(-86, "Hyperlink."),

  /// Followed Hyperlink.
  HYPERLINK_FOLLOWED(-87, "Followed Hyperlink."),

  /// Index 1.
  INDEX_1(-11, "Index 1."),

  /// Index 2.
  INDEX_2(-12, "Index 2."),

  /// Index 3.
  INDEX_3(-13, "Index 3."),

  /// Index 4.
  INDEX_4(-14, "Index 4."),

  /// Index 5.
  INDEX_5(-15, "Index 5."),

  /// Index 6.
  INDEX_6(-16, "Index 6."),

  /// Index 7.
  INDEX_7(-17, "Index 7."),

  /// Index 8.
  INDEX_8(-18, "Index 8."),

  /// Index 9.
  INDEX_9(-19, "Index 9."),

  /// Index Heading
  INDEX_HEADING(-34, "Index Heading"),

  /// Intense Emphasis.
  INTENSE_EMPHASIS(-262, "Intense Emphasis."),

  /// Intense Quote.
  INTENSE_QUOTE(-182, "Intense Quote."),

  /// Intense Reference.
  INTENSE_REFERENCE(-264, "Intense Reference."),

  /// Line Number.
  LINE_NUMBER(-41, "Line Number."),

  /// List.
  LIST(-48, "List."),

  /// List 2.
  LIST_2(-51, "List 2."),

  /// List 3.
  LIST_3(-52, "List 3."),

  /// List 4.
  LIST_4(-53, "List 4."),

  /// List 5.
  LIST_5(-54, "List 5."),

  /// List Bullet.
  LIST_BULLET(-49, "List Bullet."),

  /// List Bullet 2.
  LIST_BULLET_2(-55, "List Bullet 2."),

  /// List Bullet 3.
  LIST_BULLET_3(-56, "List Bullet 3."),

  /// List Bullet 4.
  LIST_BULLET_4(-57, "List Bullet 4."),

  /// List Bullet 5.
  LIST_BULLET_5(-58, "List Bullet 5."),

  /// List Continue.
  LIST_CONTINUE(-69, "List Continue."),

  /// List Continue 2.
  LIST_CONTINUE_2(-70, "List Continue 2."),

  /// List Continue 3.
  LIST_CONTINUE_3(-71, "List Continue 3."),

  /// List Continue 4.
  LIST_CONTINUE_4(-72, "List Continue 4."),

  /// List Continue 5.
  LIST_CONTINUE_5(-73, "List Continue 5."),

  /// List Number.
  LIST_NUMBER(-50, "List Number."),

  /// List Number 2.
  LIST_NUMBER_2(-59, "List Number 2."),

  /// List Number 3.
  LIST_NUMBER_3(-60, "List Number 3."),

  /// List Number 4.
  LIST_NUMBER_4(-61, "List Number 4."),

  /// List Number 5.
  LIST_NUMBER_5(-62, "List Number 5."),

  /// List Paragraph.
  LIST_PARAGRAPH(-180, "List Paragraph."),

  /// Macro Text.
  MACRO_TEXT(-46, "Macro Text."),

  /// Message Header.
  MESSAGE_HEADER(-74, "Message Header."),

  /// Document Map.
  NAV_PANE(-90, "Document Map."),

  /// Normal.
  NORMAL(-1, "Normal."),

  /// Normal Indent.
  NORMAL_INDENT(-29, "Normal Indent."),

  /// Normal (applied to an object).
  NORMAL_OBJECT(-158, "Normal (applied to an object)."),

  /// Normal (applied within a table).
  NORMAL_TABLE(-106, "Normal (applied within a table)."),

  /// Note Heading.
  NOTE_HEADING(-80, "Note Heading."),

  /// Page Number.
  PAGE_NUMBER(-42, "Page Number."),

  /// Plain Text.
  PLAIN_TEXT(-91, "Plain Text."),

  /// Quote.
  QUOTE(-181, "Quote."),

  /// Salutation.
  SALUTATION(-76, "Salutation."),

  /// Signature.
  SIGNATURE(-65, "Signature."),

  /// Strong.
  STRONG(-88, "Strong."),

  /// Subtitle.
  SUBTITLE(-75, "Subtitle."),

  /// Subtle Emphasis.
  SUBTLE_EMPHASIS(-261, "Subtle Emphasis."),

  /// Subtle Reference.
  SUBTLE_REFERENCE(-263, "Subtle Reference."),

  /// Colorful Grid.
  TABLE_COLORFUL_GRID(-172, "Colorful Grid."),

  /// Colorful List.
  TABLE_COLORFUL_LIST(-171, "Colorful List."),

  /// Colorful Shading.
  TABLE_COLORFUL_SHADING(-170, "Colorful Shading."),

  /// Dark List.
  TABLE_DARK_LIST(-169, "Dark List."),

  /// Light Grid.
  TABLE_LIGHT_GRID(-161, "Light Grid."),

  /// Light Grid Accent 1.
  TABLE_LIGHT_GRID_ACCENT_1(-175, "Light Grid Accent 1."),

  /// Light List.
  TABLE_LIGHT_LIST(-160, "Light List."),

  /// Light List Accent 1.
  TABLE_LIGHT_LIST_ACCENT_1(-174, "Light List Accent 1."),

  /// Light Shading.
  TABLE_LIGHT_SHADING(-159, "Light Shading."),

  /// Light Shading Accent 1.
  TABLE_LIGHT_SHADING_ACCENT_1(-173, "Light Shading Accent 1."),

  /// Medium Grid 1.
  TABLE_MEDIUM_GRID_1(-166, "Medium Grid 1."),

  /// Medium Grid 2.
  TABLE_MEDIUM_GRID_2(-167, "Medium Grid 2."),

  /// Medium Grid 3.
  TABLE_MEDIUM_GRID_3(-168, "Medium Grid 3."),

  /// Medium List 1.
  TABLE_MEDIUM_LIST_1(-164, "Medium List 1."),

  /// Medium List 1 Accent 1.
  TABLE_MEDIUM_LIST_1_ACCENT_1(-178, "Medium List 1 Accent 1."),

  /// Medium List 2.
  TABLE_MEDIUM_LIST_2(-165, "Medium List 2."),

  /// Medium Shading 1.
  TABLE_MEDIUM_SHADING_1(-162, "Medium Shading 1."),

  /// Medium Shading 1 Accent 1.
  TABLE_MEDIUM_SHADING_1_ACCENT_1(-176, "Medium Shading 1 Accent 1."),

  /// Medium Shading 2.
  TABLE_MEDIUM_SHADING_2(-163, "Medium Shading 2."),

  /// Medium Shading 2 Accent 1.
  TABLE_MEDIUM_SHADING_2_ACCENT_1(-177, "Medium Shading 2 Accent 1."),

  /// Table of Authorities.
  TABLE_OF_AUTHORITIES(-45, "Table of Authorities."),

  /// Table of Figures.
  TABLE_OF_FIGURES(-36, "Table of Figures."),

  /// Title.
  TITLE(-63, "Title."),

  /// TOA Heading.
  TOAHEADING(-47, "TOA Heading."),

  /// TOC 1.
  TOC_1(-20, "TOC 1."),

  /// TOC 2.
  TOC_2(-21, "TOC 2."),

  /// TOC 3.
  TOC_3(-22, "TOC 3."),

  /// TOC 4.
  TOC_4(-23, "TOC 4."),

  /// TOC 5.
  TOC_5(-24, "TOC 5."),

  /// TOC 6.
  TOC_6(-25, "TOC 6."),

  /// TOC 7.
  TOC_7(-26, "TOC 7."),

  /// TOC 8.
  TOC_8(-27, "TOC 8."),

  /// TOC 9.
  TOC_9(-28, "TOC 9.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_BUILTIN_STYLE(this.intValue, this.description);

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_BUILTIN_STYLE].
typedef WD_STYLE = WD_BUILTIN_STYLE;

/// Specifies one of the four style types: paragraph, character, list, or table.
///
/// Example:
/// ```dart
/// // Assuming 'style' is an object with a 'type' property
/// // assert(style.type == WD_STYLE_TYPE.PARAGRAPH);
/// ```
///
/// MS API name: `WdStyleType`
///
/// http://msdn.microsoft.com/en-us/library/office/ff196870.aspx
enum WD_STYLE_TYPE {
  /// Character style.
  CHARACTER(2, "character", "Character style."),

  /// List style.
  LIST(4, "numbering", "List style."),

  /// Paragraph style.
  PARAGRAPH(1, "paragraph", "Paragraph style."),

  /// Table style.
  TABLE(3, "table", "Table style.");

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The XML attribute value corresponding to this member.
  final String xmlValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_STYLE_TYPE(this.intValue, this.xmlValue, this.description);

  /// Returns the enum member corresponding to the XML attribute value `xmlValue`.
  static WD_STYLE_TYPE fromXml(String? xmlValue) {
    if (xmlValue == null) {
      throw ArgumentError(
          'WD_STYLE_TYPE.fromXml() requires a non-null String argument');
    }
    for (final member in values) {
      if (member.xmlValue == xmlValue) {
        return member;
      }
    }
    throw ArgumentError('WD_STYLE_TYPE has no XML mapping for "$xmlValue"');
  }

  /// Returns the XML value for the given enum member `value`.
  static String? toXml(WD_STYLE_TYPE? value) {
    return value?.xmlValue;
  }

  @override
  String toString() => '$name ($intValue)';
}

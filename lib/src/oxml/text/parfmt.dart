/// Path: lib/src/oxml/text/parfmt.dart
/// Based on python-docx: docx/oxml/text/parfmt.py
/// Custom element classes related to paragraph properties (CT_PPr).

import 'package:xml/xml.dart';

import '../../enum/text.dart'
    show WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER;
import '../../shared.dart' show Length, Twips;
import '../ns.dart' show qn; // qn and namespaces
import '../numbering.dart' show CT_NumPr; // CT_NumPr definition
import '../parser.dart' show OxmlElement; // OxmlElement
import '../section.dart' show CT_SectPr; // CT_SectPr definition
// Assume CT_String and CT_OnOff are defined here and CT_OnOff has create(qn)
import '../shared.dart' show CT_OnOff, CT_String;
// Assume converters are defined here or imported elsewhere
import '../simpletypes.dart';
import '../xmlchemy.dart' show BaseOxmlElement; // BaseOxmlElement
/// Path: lib/src/oxml/text/parfmt.dart
/// Based on python-docx: docx/oxml/text/parfmt.py
/// Custom element classes related to paragraph properties (CT_PPr).

// Assuming these are defined/imported from simpletypes.dart or similar
class ST_SignedTwipsMeasureConverter implements BaseSimpleType<Length> {
  const ST_SignedTwipsMeasureConverter(); // Make const if possible
  @override
  Length fromXml(String xmlValue) => Twips(int.parse(xmlValue));
  @override
  String? toXml(Length? value) => value?.twips.toString();
  @override
  void validate(Length value) {} // Add range checks if necessary
}

class ST_TwipsMeasureConverter implements BaseSimpleType<Length> {
  const ST_TwipsMeasureConverter(); // Make const if possible
  @override
  Length fromXml(String xmlValue) => Twips(int.parse(xmlValue));
  @override
  String? toXml(Length? value) => value?.twips.toString();
  @override
  void validate(Length value) {} // Add range checks if necessary
}

// Assuming your Enums have an 'xmlValue' getter
class WD_ALIGN_PARAGRAPH_Converter
    implements BaseSimpleType<WD_ALIGN_PARAGRAPH> {
  const WD_ALIGN_PARAGRAPH_Converter(); // Make const if possible
  @override
  WD_ALIGN_PARAGRAPH fromXml(String xmlValue) =>
      WD_ALIGN_PARAGRAPH.values.firstWhere((e) => e.xmlValue == xmlValue,
          orElse: () => throw FormatException(
              "Invalid WD_ALIGN_PARAGRAPH value: $xmlValue"));
  @override
  String? toXml(WD_ALIGN_PARAGRAPH? value) => value?.xmlValue;
  @override
  void validate(WD_ALIGN_PARAGRAPH value) {}
}

class WD_LINE_SPACING_Converter implements BaseSimpleType<WD_LINE_SPACING> {
  const WD_LINE_SPACING_Converter(); // Make const if possible
  @override
  WD_LINE_SPACING fromXml(String xmlValue) => WD_LINE_SPACING.values.firstWhere(
      (e) => e.xmlValue == xmlValue,
      orElse: () =>
          throw FormatException("Invalid WD_LINE_SPACING value: $xmlValue"));
  @override
  String? toXml(WD_LINE_SPACING? value) => value?.xmlValue;
  @override
  void validate(WD_LINE_SPACING value) {}
}

class WD_TAB_ALIGNMENT_Converter implements BaseSimpleType<WD_TAB_ALIGNMENT> {
  const WD_TAB_ALIGNMENT_Converter(); // Make const if possible
  @override
  WD_TAB_ALIGNMENT fromXml(String xmlValue) =>
      WD_TAB_ALIGNMENT.values.firstWhere((e) => e.xmlValue == xmlValue,
          orElse: () => throw FormatException(
              "Invalid WD_TAB_ALIGNMENT value: $xmlValue"));
  @override
  String? toXml(WD_TAB_ALIGNMENT? value) => value?.xmlValue;
  @override
  void validate(WD_TAB_ALIGNMENT value) {}
}

class WD_TAB_LEADER_Converter implements BaseSimpleType<WD_TAB_LEADER> {
  const WD_TAB_LEADER_Converter(); // Make const if possible
  @override
  WD_TAB_LEADER fromXml(String xmlValue) =>
      WD_TAB_LEADER.values.firstWhere((e) => e.xmlValue == xmlValue,
          orElse: () =>
              throw FormatException("Invalid WD_TAB_LEADER value: $xmlValue"));
  @override
  String? toXml(WD_TAB_LEADER? value) => value?.xmlValue;
  @override
  void validate(WD_TAB_LEADER value) {}
}

// Instantiate the converters
final stSignedTwipsMeasureConverter = const ST_SignedTwipsMeasureConverter();
final stTwipsMeasureConverter = const ST_TwipsMeasureConverter();
final wdAlignParagraphConverter = const WD_ALIGN_PARAGRAPH_Converter();
final wdLineSpacingConverter = const WD_LINE_SPACING_Converter();
final wdTabAlignmentConverter = const WD_TAB_ALIGNMENT_Converter();
final wdTabLeaderConverter = const WD_TAB_LEADER_Converter();

// --- End of Moved Converters ---

/// `<w:ind>` element, specifying paragraph indentation.
class CT_Ind extends BaseOxmlElement {
  CT_Ind(super.element);

  /// Creates a new `<w:ind>` element.
  static XmlElement create() => OxmlElement(qnTagName);

  /// Indentation from the left edge.
  Length? get left => getAttrVal('w:left', stSignedTwipsMeasureConverter);
  set left(Length? value) =>
      setAttrVal('w:left', value, stSignedTwipsMeasureConverter);

  /// Indentation from the right edge.
  Length? get right => getAttrVal('w:right', stSignedTwipsMeasureConverter);
  set right(Length? value) =>
      setAttrVal('w:right', value, stSignedTwipsMeasureConverter);

  /// Additional indentation for the first line.
  Length? get firstLine => getAttrVal('w:firstLine', stTwipsMeasureConverter);
  set firstLine(Length? value) =>
      setAttrVal('w:firstLine', value, stTwipsMeasureConverter);

  /// Indentation for all lines except the first (hanging indent).
  Length? get hanging => getAttrVal('w:hanging', stTwipsMeasureConverter);
  set hanging(Length? value) =>
      setAttrVal('w:hanging', value, stTwipsMeasureConverter);

  static final qnTagName = qn('w:ind');
}

/// `<w:jc>` element, specifying paragraph justification.
class CT_Jc extends BaseOxmlElement {
  CT_Jc(super.element);

  /// Creates a new `<w:jc>` element with the specified alignment.
  static XmlElement create(WD_ALIGN_PARAGRAPH alignment) {
    final xmlVal = wdAlignParagraphConverter.toXml(alignment);
    if (xmlVal == null) {
      throw ArgumentError("Cannot create CT_Jc with null alignment value");
    }
    return OxmlElement(qnTagName, attrs: {qn('w:val'): xmlVal});
  }

  /// The alignment value (e.g., LEFT, CENTER).
  WD_ALIGN_PARAGRAPH get val =>
      getReqAttrVal('w:val', wdAlignParagraphConverter);
  set val(WD_ALIGN_PARAGRAPH value) =>
      setReqAttrVal('w:val', value, wdAlignParagraphConverter);

  static final qnTagName = qn('w:jc');
}

/// `<w:pPr>` element, containing the properties for a paragraph.
class CT_PPr extends BaseOxmlElement {
  CT_PPr(super.element);

  /// Creates a new `<w:pPr>` element.
  static XmlElement create() => OxmlElement(qnTagName);

  // Define the sequence of child elements for insertion logic
  // Use qualified names directly for clarity and robustness
  static final List<String> _tagSeq = [
    qn("w:pStyle"),
    qn("w:keepNext"),
    qn("w:keepLines"),
    qn("w:pageBreakBefore"),
    qn("w:framePr"),
    qn("w:widowControl"),
    qn("w:numPr"),
    qn("w:suppressLineNumbers"),
    qn("w:pBdr"),
    qn("w:shd"),
    qn("w:tabs"),
    qn("w:suppressAutoHyphens"),
    qn("w:kinsoku"),
    qn("w:wordWrap"),
    qn("w:overflowPunct"),
    qn("w:topLinePunct"),
    qn("w:autoSpaceDE"),
    qn("w:autoSpaceDN"),
    qn("w:bidi"),
    qn("w:adjustRightInd"),
    qn("w:snapToGrid"),
    qn("w:spacing"),
    qn("w:ind"),
    qn("w:contextualSpacing"),
    qn("w:mirrorIndents"),
    qn("w:suppressOverlap"),
    qn("w:jc"),
    qn("w:textDirection"),
    qn("w:textAlignment"),
    qn("w:textboxTightWrap"),
    qn("w:outlineLvl"),
    qn("w:divId"),
    qn("w:cnfStyle"),
    qn("w:rPr"),
    qn("w:sectPr"),
    qn("w:pPrChange"),
  ];

  // --- Getters for ZeroOrOne child elements (using qualified names) ---
  CT_String? get pStyle => _getCTString(_tagSeq[0]);
  CT_OnOff? get keepNext => _getCTOnOff(_tagSeq[1]);
  CT_OnOff? get keepLines => _getCTOnOff(_tagSeq[2]);
  CT_OnOff? get pageBreakBefore => _getCTOnOff(_tagSeq[3]);
  CT_OnOff? get widowControl => _getCTOnOff(_tagSeq[5]);
  CT_NumPr? get numPr => _getCTNumPr(_tagSeq[6]);
  CT_TabStops? get tabs => _getCTTabStops(_tagSeq[10]);
  CT_Spacing? get spacing => _getCTSpacing(_tagSeq[21]);
  CT_Ind? get ind => _getCTInd(_tagSeq[22]);
  CT_Jc? get jc => _getCTJc(_tagSeq[26]);
  CT_SectPr? get sectPr => _getCTSectPr(_tagSeq[34]);

  // --- Helper Getters for Child Wrappers ---
  CT_String? _getCTString(String tag) =>
      childOrNull(tag) == null ? null : CT_String(childOrNull(tag)!);
  CT_OnOff? _getCTOnOff(String tag) =>
      childOrNull(tag) == null ? null : CT_OnOff(childOrNull(tag)!);
  CT_NumPr? _getCTNumPr(String tag) =>
      childOrNull(tag) == null ? null : CT_NumPr(childOrNull(tag)!);
  CT_TabStops? _getCTTabStops(String tag) =>
      childOrNull(tag) == null ? null : CT_TabStops(childOrNull(tag)!);
  CT_Spacing? _getCTSpacing(String tag) =>
      childOrNull(tag) == null ? null : CT_Spacing(childOrNull(tag)!);
  CT_Ind? _getCTInd(String tag) =>
      childOrNull(tag) == null ? null : CT_Ind(childOrNull(tag)!);
  CT_Jc? _getCTJc(String tag) =>
      childOrNull(tag) == null ? null : CT_Jc(childOrNull(tag)!);
  CT_SectPr? _getCTSectPr(String tag) =>
      childOrNull(tag) == null ? null : CT_SectPr(childOrNull(tag)!);

  // --- Get or Add Methods (using qualified names) ---
  CT_Ind getOrAddInd() =>
      CT_Ind(getOrAddChild(_tagSeq[22], _tagSeq.sublist(23), CT_Ind.create));
  CT_String getOrAddPStyle() => CT_String(getOrAddChild(
      _tagSeq[0], _tagSeq.sublist(1), () => CT_String.create(_tagSeq[0], '')));
  CT_Jc getOrAddJc() => CT_Jc(getOrAddChild(_tagSeq[26], _tagSeq.sublist(27),
      () => CT_Jc.create(WD_ALIGN_PARAGRAPH.LEFT)));
  CT_Spacing getOrAddSpacing() => CT_Spacing(
      getOrAddChild(_tagSeq[21], _tagSeq.sublist(22), CT_Spacing.create));
  CT_TabStops getOrAddTabs() => CT_TabStops(
      getOrAddChild(_tagSeq[10], _tagSeq.sublist(11), CT_TabStops.create));

  // --- Corrected Get or Add Methods for CT_OnOff ---
  // Assumes CT_OnOff.create(qnTagName) exists and creates an empty element <tagName/>
  CT_OnOff getOrAddKeepNext() => CT_OnOff(getOrAddChild(
      _tagSeq[1], _tagSeq.sublist(2), () => CT_OnOff.create(_tagSeq[1])));
  CT_OnOff getOrAddKeepLines() => CT_OnOff(getOrAddChild(
      _tagSeq[2], _tagSeq.sublist(3), () => CT_OnOff.create(_tagSeq[2])));
  CT_OnOff getOrAddPageBreakBefore() => CT_OnOff(getOrAddChild(
      _tagSeq[3], _tagSeq.sublist(4), () => CT_OnOff.create(_tagSeq[3])));
  CT_OnOff getOrAddWidowControl() => CT_OnOff(getOrAddChild(
      _tagSeq[5], _tagSeq.sublist(6), () => CT_OnOff.create(_tagSeq[5])));

  // --- Remove Methods (using qualified names) ---
  void removeInd() => removeChild(_tagSeq[22]);
  void removePStyle() => removeChild(_tagSeq[0]);
  void removeJc() => removeChild(_tagSeq[26]);
  void removeSpacing() => removeChild(_tagSeq[21]);
  void removeTabs() => removeChild(_tagSeq[10]);
  void removeKeepNext() => removeChild(_tagSeq[1]);
  void removeKeepLines() => removeChild(_tagSeq[2]);
  void removePageBreakBefore() => removeChild(_tagSeq[3]);
  void removeWidowControl() => removeChild(_tagSeq[5]);
  void removeSectPr() => removeChild(_tagSeq[34]);

  // --- Special Properties Combining Children/Attributes ---

  /// Combined first line/hanging indent value.
  Length? get firstLineIndent {
    final indElement = ind;
    if (indElement == null) return null;
    final hanging = indElement.hanging;
    if (hanging != null)
      return Length(-hanging.emu); // Negative value for hanging
    return indElement.firstLine; // Returns null if firstLine is also null
  }

  set firstLineIndent(Length? value) {
    // Ensure <w:ind> exists if we need to set or clear values
    if (ind == null && value == null) return;
    final indElement = getOrAddInd();
    // Clear existing settings first
    indElement.firstLine = null;
    indElement.hanging = null;

    if (value == null) {
      // If setting to null and no other indent attrs exist, remove <w:ind>
      if (indElement.left == null && indElement.right == null) {
        removeInd();
      }
      return;
    }
    if (value.emu < 0) {
      indElement.hanging = Length(-value.emu);
    } else {
      indElement.firstLine = value;
    }
  }

  /// Left indentation value.
  Length? get indLeft => ind?.left;
  set indLeft(Length? value) {
    if (ind == null && value == null) return;
    final indElement = getOrAddInd();
    indElement.left = value;
    // If clearing and no other indent attrs exist, remove <w:ind>
    if (value == null &&
        indElement.right == null &&
        indElement.firstLine == null &&
        indElement.hanging == null) {
      removeInd();
    }
  }

  /// Right indentation value.
  Length? get indRight => ind?.right;
  set indRight(Length? value) {
    if (ind == null && value == null) return;
    final indElement = getOrAddInd();
    indElement.right = value;
    // If clearing and no other indent attrs exist, remove <w:ind>
    if (value == null &&
        indElement.left == null &&
        indElement.firstLine == null &&
        indElement.hanging == null) {
      removeInd();
    }
  }

  /// Paragraph alignment/justification value.
  WD_ALIGN_PARAGRAPH? get jcVal => jc?.val;
  set jcVal(WD_ALIGN_PARAGRAPH? value) {
    if (value == null) {
      removeJc();
    } else {
      getOrAddJc().val = value;
    }
  }

  /// Keep lines together value.
  bool? get keepLinesVal => keepLines?.val;
  set keepLinesVal(bool? value) {
    if (value == null) {
      removeKeepLines();
    } else {
      getOrAddKeepLines().val = value;
    }
  }

  /// Keep with next value.
  bool? get keepNextVal => keepNext?.val;
  set keepNextVal(bool? value) {
    if (value == null) {
      removeKeepNext();
    } else {
      getOrAddKeepNext().val = value;
    }
  }

  /// Page break before value.
  bool? get pageBreakBeforeVal => pageBreakBefore?.val;
  set pageBreakBeforeVal(bool? value) {
    if (value == null) {
      removePageBreakBefore();
    } else {
      getOrAddPageBreakBefore().val = value;
    }
  }

  /// Spacing after paragraph.
  Length? get spacingAfter => spacing?.after;
  set spacingAfter(Length? value) {
    if (spacing == null && value == null) return;
    final spacingElement = getOrAddSpacing();
    spacingElement.after = value;
    // If clearing and no other spacing attrs exist, remove <w:spacing>
    if (value == null &&
        spacingElement.before == null &&
        spacingElement.line == null &&
        spacingElement.lineRule == null) {
      removeSpacing();
    }
  }

  /// Spacing before paragraph.
  Length? get spacingBefore => spacing?.before;
  set spacingBefore(Length? value) {
    if (spacing == null && value == null) return;
    final spacingElement = getOrAddSpacing();
    spacingElement.before = value;
    if (value == null &&
        spacingElement.after == null &&
        spacingElement.line == null &&
        spacingElement.lineRule == null) {
      removeSpacing();
    }
  }

  /// Line spacing value (either multiple as `double` or exact/atLeast as `Length`).
  dynamic get spacingLine {
    // Returns double or Length?
    final spacingElement = spacing;
    if (spacingElement?.line == null) return null;
    // Default rule is MULTIPLE if line is set but rule isn't explicitly 'exact' or 'atLeast'
    final lineRule = spacingElement?.lineRule ?? WD_LINE_SPACING.MULTIPLE;
    final lineVal = spacingElement!.line!;

    if (lineRule == WD_LINE_SPACING.MULTIPLE) {
      // In OOXML, multiple is stored in units of 240ths of a line.
      // 240 = single, 360 = 1.5, 480 = double
      return lineVal.emu / Twips(240).emu; // Return as double multiple
    } else {
      // AT_LEAST or EXACTLY
      return lineVal; // Return as Length
    }
  }

  set spacingLine(dynamic value) {
    // Accepts double or Length?
    if (spacing == null && value == null) return;
    final spacingElement = getOrAddSpacing();

    if (value == null) {
      spacingElement.line = null;
      // Also clear the rule when clearing the line value? Seems logical.
      spacingElement.lineRule = null;
      // Clean up <w:spacing> if empty
      if (spacingElement.after == null && spacingElement.before == null) {
        removeSpacing();
      }
      return;
    }

    if (value is num) {
      // Multiple line spacing
      spacingElement.line = Twips((value * 240).round());
      spacingElement.lineRule = WD_LINE_SPACING.MULTIPLE;
    } else if (value is Length) {
      // Exact or AtLeast spacing
      spacingElement.line = value;
      // Preserve AT_LEAST if already set, otherwise default to EXACTLY
      if (spacingElement.lineRule != WD_LINE_SPACING.AT_LEAST) {
        spacingElement.lineRule = WD_LINE_SPACING.EXACTLY;
      }
    } else {
      throw ArgumentError(
          "line_spacing must be double, int or Length object, got ${value.runtimeType}");
    }
  }

  /// Line spacing rule (SINGLE, MULTIPLE, AT_LEAST, etc.).
  WD_LINE_SPACING? get spacingLineRule {
    final spacingElement = spacing;
    if (spacingElement == null) return null;
    final line = spacingElement.line;
    final rule = spacingElement.lineRule;

    // Infer rule if not explicitly set but line value exists
    if (rule == null && line != null) {
      return WD_LINE_SPACING
          .MULTIPLE; // Default interpretation if line is set without rule
    }
    // Check for special MULTIPLE cases (single, 1.5, double)
    if (rule == WD_LINE_SPACING.MULTIPLE && line != null) {
      if (line.twips == 240) return WD_LINE_SPACING.SINGLE;
      if (line.twips == 360) return WD_LINE_SPACING.ONE_POINT_FIVE;
      if (line.twips == 480) return WD_LINE_SPACING.DOUBLE;
      // If it's MULTIPLE but not one of the standard ones, return MULTIPLE
      return WD_LINE_SPACING.MULTIPLE;
    }
    return rule; // Return EXACTLY, AT_LEAST, or null if rule is explicitly null
  }

  set spacingLineRule(WD_LINE_SPACING? value) {
    if (spacing == null && value == null) return;
    final spacingElement = getOrAddSpacing();

    if (value == WD_LINE_SPACING.SINGLE) {
      spacingElement.line = Twips(240);
      spacingElement.lineRule = WD_LINE_SPACING.MULTIPLE;
    } else if (value == WD_LINE_SPACING.ONE_POINT_FIVE) {
      spacingElement.line = Twips(360);
      spacingElement.lineRule = WD_LINE_SPACING.MULTIPLE;
    } else if (value == WD_LINE_SPACING.DOUBLE) {
      spacingElement.line = Twips(480);
      spacingElement.lineRule = WD_LINE_SPACING.MULTIPLE;
    } else {
      // For EXACTLY, AT_LEAST, or null (to clear)
      spacingElement.lineRule = value;
      // If rule is set to EXACTLY or AT_LEAST without a line value,
      // should we add a default line value (e.g., 12pt)? Python didn't.
      // Let's clear line value only if value is null to mirror Python.
      if (value == null) {
        spacingElement.line = null;
      }
    }
    // Clean up <w:spacing> if empty
    if (spacingElement.after == null &&
        spacingElement.before == null &&
        spacingElement.line == null &&
        spacingElement.lineRule == null) {
      removeSpacing();
    }
  }

  /// Paragraph style ID.
  String? get style => pStyle?.val;
  set style(String? value) {
    if (value == null) {
      removePStyle();
    } else {
      getOrAddPStyle().val = value;
    }
  }

  /// Widow/orphan control value.
  bool? get widowControlVal => widowControl?.val;
  set widowControlVal(bool? value) {
    if (value == null) {
      removeWidowControl();
    } else {
      getOrAddWidowControl().val = value;
    }
  }

  // --- Methods for managing sectPr ---
  void insertSectPr(CT_SectPr sectPr) {
    // sectPr is the last element according to _tagSeq
    // Use the insertChild method from BaseOxmlElement
    insertChild(sectPr.element, _tagSeq.sublist(35));
  }
  // removeSectPr() already implemented via removeChild above

  static final qnTagName = qn('w:pPr');
}

/// `<w:spacing>` element, specifying paragraph spacing attributes.
class CT_Spacing extends BaseOxmlElement {
  CT_Spacing(super.element);

  /// Creates a new `<w:spacing>` element.
  static XmlElement create() => OxmlElement(qnTagName);

  /// Spacing after the paragraph.
  Length? get after => getAttrVal('w:after', stTwipsMeasureConverter);
  set after(Length? value) =>
      setAttrVal('w:after', value, stTwipsMeasureConverter);

  /// Spacing before the paragraph.
  Length? get before => getAttrVal('w:before', stTwipsMeasureConverter);
  set before(Length? value) =>
      setAttrVal('w:before', value, stTwipsMeasureConverter);

  /// Line spacing value (depends on lineRule). Represented as twips internally.
  Length? get line => getAttrVal('w:line', stSignedTwipsMeasureConverter);
  set line(Length? value) =>
      setAttrVal('w:line', value, stSignedTwipsMeasureConverter);

  /// Line spacing rule (e.g., auto, exact, atLeast).
  WD_LINE_SPACING? get lineRule => getAttrVal(
      'w:lineRule', wdLineSpacingConverter); // Use converter instance
  set lineRule(WD_LINE_SPACING? value) => setAttrVal(
      'w:lineRule', value, wdLineSpacingConverter); // Use converter instance

  static final qnTagName = qn('w:spacing');
}

/// `<w:tab>` element, representing an individual tab stop defined in `<w:tabs>`.
/// Also used for the `<w:tab/>` character in runs, hence the `toString()`.
class CT_TabStop extends BaseOxmlElement {
  CT_TabStop(super.element);

  /// Creates a new `<w:tab>` element for defining a tab stop.
  static XmlElement create({
    required Length pos,
    required WD_TAB_ALIGNMENT align,
    WD_TAB_LEADER? leader,
  }) {
    final attrs = <String, String>{
      qn('w:pos'): stSignedTwipsMeasureConverter.toXml(pos)!,
      qn('w:val'):
          wdTabAlignmentConverter.toXml(align)!, // Use converter instance
    };
    // Only add leader attribute if it's not the default (SPACES)
    final leaderXml =
        wdTabLeaderConverter.toXml(leader); // Use converter instance
    if (leaderXml != null && leader != WD_TAB_LEADER.SPACES) {
      attrs[qn('w:leader')] = leaderXml;
    }
    return OxmlElement(qnTagName, attrs: attrs);
  }

  /// Creates a new `<w:tab/>` element representing a tab character in a run.
  static XmlElement createRunTab() =>
      OxmlElement(qnTagName); // Empty element for runs

  /// Tab stop alignment.
  WD_TAB_ALIGNMENT get val =>
      getReqAttrVal('w:val', wdTabAlignmentConverter); // Use converter instance
  set val(WD_TAB_ALIGNMENT value) => setReqAttrVal(
      'w:val', value, wdTabAlignmentConverter); // Use converter instance

  /// Tab stop leader character. Defaults to SPACES if attribute not present.
  WD_TAB_LEADER get leader => getAttrVal('w:leader', wdTabLeaderConverter,
      defaultValue: WD_TAB_LEADER.SPACES)!; // Use converter instance
  set leader(WD_TAB_LEADER? value) => // Allow null to remove/set default
      setAttrVal('w:leader', value, wdTabLeaderConverter,
          defaultValue: WD_TAB_LEADER.SPACES); // Use converter instance

  /// Tab stop position relative to the paragraph edge.
  Length get pos => getReqAttrVal('w:pos', stSignedTwipsMeasureConverter);
  set pos(Length value) =>
      setReqAttrVal('w:pos', value, stSignedTwipsMeasureConverter);

  /// Text equivalent of a `<w:tab/>` element appearing in a run.
  @override
  String toString() => "\t";

  static final qnTagName = qn('w:tab');
}

/// `<w:tabs>` element, container for a sorted sequence of tab stops.
class CT_TabStops extends BaseOxmlElement {
  CT_TabStops(super.element);

  /// Creates a new `<w:tabs>` element.
  static XmlElement create() => OxmlElement(qnTagName);

  /// List of `<w:tab>` child elements defining individual tab stops.
  List<CT_TabStop> get tabElements => childrenWhereType<CT_TabStop>(
      CT_TabStop.qnTagName, (el) => CT_TabStop(el));

  /// Alias for `tabElements` to match Python's `tab_lst`.
  List<CT_TabStop> get tab_lst => tabElements;

  /// Helper to create a new underlying tab stop element.
  CT_TabStop _newTab({
    required Length pos,
    required WD_TAB_ALIGNMENT align,
    WD_TAB_LEADER? leader,
  }) {
    // Use the static create method of CT_TabStop
    return CT_TabStop(
        CT_TabStop.create(pos: pos, align: align, leader: leader));
  }

  /// Insert a newly created `w:tab` child element in `pos` order.
  /// Returns the wrapper for the newly inserted tab stop.
  CT_TabStop insertTabInOrder(
      Length pos, WD_TAB_ALIGNMENT align, WD_TAB_LEADER leader) {
    final newTab = _newTab(pos: pos, align: align, leader: leader);

    // Find the correct insertion index based on position
    int insertIdx = 0;
    for (final existingTab in tabElements) {
      if (pos.emu < existingTab.pos.emu) {
        break; // Found the first tab with a greater position
      }
      insertIdx++;
    }

    // Insert the actual XmlElement into the parent's children list
    element.children.insert(insertIdx, newTab.element);
    return newTab; // Return the wrapper
  }

  static final qnTagName = qn('w:tabs');
}

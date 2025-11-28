import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';
import 'package:docx_dart/src/oxml/text/parfmt.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/text/tabstops.dart';

/// Provides access to paragraph-level formatting options.
class ParagraphFormat extends ElementProxy {
  ParagraphFormat(CT_P paragraph)
      : _paragraph = paragraph,
        super(paragraph);

  final CT_P _paragraph;
  TabStops? _tabStops;

  CT_PPr? get _pPr => _paragraph.pPr;
  CT_PPr get _pPrRequired => _paragraph.getOrAddPPr();

  WD_PARAGRAPH_ALIGNMENT? get alignment => _pPr?.jcVal;
  set alignment(WD_PARAGRAPH_ALIGNMENT? value) => _pPrRequired.jcVal = value;

  Length? get firstLineIndent => _pPr?.firstLineIndent;
  set firstLineIndent(Length? value) => _pPrRequired.firstLineIndent = value;

  bool? get keepTogether => _pPr?.keepLinesVal;
  set keepTogether(bool? value) => _pPrRequired.keepLinesVal = value;

  bool? get keepWithNext => _pPr?.keepNextVal;
  set keepWithNext(bool? value) => _pPrRequired.keepNextVal = value;

  Length? get leftIndent => _pPr?.indLeft;
  set leftIndent(Length? value) => _pPrRequired.indLeft = value;

  dynamic get lineSpacing {
    final spacing = _pPr?.spacing;
    if (spacing == null || spacing.line == null) {
      return null;
    }
    final rule = spacing.lineRule ?? WD_LINE_SPACING.MULTIPLE;
    final line = spacing.line!;
    if (rule == WD_LINE_SPACING.MULTIPLE) {
      return line.twips / 240;
    }
    return line;
  }
  set lineSpacing(dynamic value) {
    final pPr = _pPrRequired;
    if (value == null) {
      pPr.spacingLine = null;
      pPr.spacingLineRule = null;
    } else if (value is Length) {
      pPr.spacingLine = value;
      if (pPr.spacingLineRule != WD_LINE_SPACING.AT_LEAST) {
        pPr.spacingLineRule = WD_LINE_SPACING.EXACTLY;
      }
    } else if (value is num) {
      pPr.spacingLine = Twips((value * 240).round());
      pPr.spacingLineRule = WD_LINE_SPACING.MULTIPLE;
    } else {
      throw ArgumentError(
          'lineSpacing must be null, Length, or num. Found ${value.runtimeType}');
    }
  }

  WD_LINE_SPACING? get lineSpacingRule {
    final spacing = _pPr?.spacing;
    if (spacing == null) {
      return null;
    }
    final line = spacing.line;
    final rule = spacing.lineRule;
    if (line == null) {
      return rule;
    }
    final effectiveRule = rule ?? WD_LINE_SPACING.MULTIPLE;
    if (effectiveRule == WD_LINE_SPACING.MULTIPLE) {
      if (line.twips == 240) return WD_LINE_SPACING.SINGLE;
      if (line.twips == 360) return WD_LINE_SPACING.ONE_POINT_FIVE;
      if (line.twips == 480) return WD_LINE_SPACING.DOUBLE;
      return WD_LINE_SPACING.MULTIPLE;
    }
    return rule;
  }
  set lineSpacingRule(WD_LINE_SPACING? value) {
    final pPr = _pPrRequired;
    if (value == WD_LINE_SPACING.SINGLE) {
      pPr.spacingLine = Twips(240);
      pPr.spacingLineRule = WD_LINE_SPACING.MULTIPLE;
    } else if (value == WD_LINE_SPACING.ONE_POINT_FIVE) {
      pPr.spacingLine = Twips(360);
      pPr.spacingLineRule = WD_LINE_SPACING.MULTIPLE;
    } else if (value == WD_LINE_SPACING.DOUBLE) {
      pPr.spacingLine = Twips(480);
      pPr.spacingLineRule = WD_LINE_SPACING.MULTIPLE;
    } else {
      pPr.spacingLineRule = value;
      if (value == null) {
        pPr.spacingLine = null;
      }
    }
  }

  bool? get pageBreakBefore => _pPr?.pageBreakBeforeVal;
  set pageBreakBefore(bool? value) => _pPrRequired.pageBreakBeforeVal = value;

  Length? get rightIndent => _pPr?.indRight;
  set rightIndent(Length? value) => _pPrRequired.indRight = value;

  Length? get spaceAfter => _pPr?.spacingAfter;
  set spaceAfter(Length? value) => _pPrRequired.spacingAfter = value;

  Length? get spaceBefore => _pPr?.spacingBefore;
  set spaceBefore(Length? value) => _pPrRequired.spacingBefore = value;

  TabStops get tabStops =>
      _tabStops ??= TabStops(_paragraph.getOrAddPPr());

  bool? get widowControl => _pPr?.widowControlVal;
  set widowControl(bool? value) => _pPrRequired.widowControlVal = value;
}

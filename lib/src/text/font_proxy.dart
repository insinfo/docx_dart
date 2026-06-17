import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/oxml/text/font.dart';
import 'package:docx_dart/src/oxml/text/run.dart';
import 'package:docx_dart/src/oxml/styles.dart' show CT_Style;
import 'package:docx_dart/src/oxml/xmlchemy.dart' show BaseOxmlElement;
import 'package:docx_dart/src/shared.dart';

/// Proxy object for character-level formatting (font) of a [Run] or [Style].
///
/// Provides convenient access to run properties such as bold, italic,
/// font name, font size, underline, color, etc.
///
/// Usage example:
/// ```dart
/// final run = paragraph.addRun('Hello');
/// run.font.bold = true;
/// run.font.size = Pt(12);
/// run.font.name = 'Arial';
/// run.font.underline = WD_UNDERLINE.SINGLE;
/// ```
class Font {
  final BaseOxmlElement _element;

  Font(this._element);

  CT_RPr get _rPr {
    final el = _element;
    if (el is CT_R) {
      return el.getOrAddRPr();
    } else if (el is CT_Style) {
      return el.getOrAddRPr();
    }
    // Generic fallback
    var rPr = _element.childOrNull(CT_RPr.qnTagName);
    if (rPr == null) {
      rPr = CT_RPr.create();
      _element.element.children.insert(0, rPr);
    }
    return CT_RPr(rPr);
  }

  CT_RPr? get _rPrOrNull {
    final el = _element;
    if (el is CT_R) {
      return el.rPr;
    } else if (el is CT_Style) {
      return el.rPrElement;
    }
    final rPr = _element.childOrNull(CT_RPr.qnTagName);
    return rPr != null ? CT_RPr(rPr) : null;
  }

  /// Whether this run is bold. `null` means inherited from style.
  bool? get bold => _rPrOrNull?.bold;
  set bold(bool? value) => _rPr.bold = value;

  /// Whether this run is italic. `null` means inherited from style.
  bool? get italic => _rPrOrNull?.italic;
  set italic(bool? value) => _rPr.italic = value;

  /// Font size as a [Length]. Use `Pt(12)` to set 12-point size.
  /// `null` means inherited from style.
  Length? get size => _rPrOrNull?.szVal;
  set size(Length? value) => _rPr.szVal = value;

  /// Font name (ASCII typeface). `null` means inherited from style.
  String? get name => _rPrOrNull?.rFontsAscii;
  set name(String? value) {
    _rPr.rFontsAscii = value;
    // Also set hAnsi to keep consistency
    _rPr.rFontsHAnsi = value;
  }

  /// Underline style. `null` means inherited from style.
  WD_UNDERLINE? get underline => _rPrOrNull?.uVal;
  set underline(WD_UNDERLINE? value) => _rPr.uVal = value;

  /// Font color as an [RGBColor]. `null` means inherited or auto.
  RGBColor? get color {
    final colorElement = _rPrOrNull?.color;
    if (colorElement == null) return null;
    final val = colorElement.val;
    return val is RGBColor ? val : null;
  }

  set color(RGBColor? value) {
    if (value == null) {
      _rPrOrNull?.removeColor();
    } else {
      _rPr.getOrAddColor().val = value;
    }
  }

  /// Whether all caps are applied. `null` means inherited from style.
  bool? get allCaps => _rPrOrNull?.caps;
  set allCaps(bool? value) => _rPr.caps = value;

  /// Whether small caps are applied. `null` means inherited from style.
  bool? get smallCaps => _rPrOrNull?.smallCaps;
  set smallCaps(bool? value) => _rPr.smallCaps = value;

  /// Whether strikethrough is applied. `null` means inherited from style.
  bool? get strike => _rPrOrNull?.strike;
  set strike(bool? value) => _rPr.strike = value;

  /// Whether double-strikethrough is applied. `null` means inherited.
  bool? get doubleStrike => _rPrOrNull?.doubleStrike;
  set doubleStrike(bool? value) => _rPr.doubleStrike = value;

  /// Whether the text is hidden. `null` means inherited from style.
  bool? get hidden => _rPrOrNull?.vanish;
  set hidden(bool? value) => _rPr.vanish = value;

  /// Whether subscript formatting is applied. `null` means inherited.
  bool? get subscript => _rPrOrNull?.subscript;
  set subscript(bool? value) => _rPr.subscript = value;

  /// Whether superscript formatting is applied. `null` means inherited.
  bool? get superscript => _rPrOrNull?.superscript;
  set superscript(bool? value) => _rPr.superscript = value;

  /// Highlight color. `null` means no highlight.
  WD_COLOR_INDEX? get highlight => _rPrOrNull?.highlightVal;
  set highlight(WD_COLOR_INDEX? value) => _rPr.highlightVal = value;
}

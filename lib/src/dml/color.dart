// docx/dml/color.dart
import 'package:docx_dart/src/enum/dml.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/oxml/text/font.dart';
import 'package:docx_dart/src/oxml/text/run.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart';

/// Provides access to run color settings (RGB, theme, automatic).
class ColorFormat extends ElementProxy {
  ColorFormat(BaseOxmlElement rParent) : super(rParent);

  /// RGB color applied directly to the run, or null when not specified.
  RGBColor? get rgb {
    final color = _color;
    if (color == null) {
      return null;
    }
    final value = color.val;
    if (value == ST_HexColorAuto.AUTO) {
      return null;
    }
    return value is RGBColor ? value : null;
  }

  set rgb(RGBColor? value) {
    final currentColor = _color;
    if (value == null && currentColor == null) {
      return;
    }
    final rPr = _r.getOrAddRPr();
    rPr.removeColor();
    if (value != null) {
      final color = rPr.getOrAddColor();
      color.val = value;
      color.themeColor = null;
    }
  }

  /// Theme color applied to the run, or null when absent.
  MSO_THEME_COLOR_INDEX? get themeColor {
    final color = _color;
    if (color == null) {
      return null;
    }
    return color.themeColor;
  }

  set themeColor(MSO_THEME_COLOR_INDEX? value) {
    if (value == null) {
      final color = _color;
      if (color != null) {
        _r.getOrAddRPr().removeColor();
      }
      return;
    }
    final color = _r.getOrAddRPr().getOrAddColor();
    color.themeColor = value;
  }

  /// How the color is defined (RGB, THEME, AUTO) or null when inherited.
  MSO_COLOR_TYPE? get type {
    final color = _color;
    if (color == null) {
      return null;
    }
    if (color.themeColor != null) {
      return MSO_COLOR_TYPE.THEME;
    }
    if (color.val == ST_HexColorAuto.AUTO) {
      return MSO_COLOR_TYPE.AUTO;
    }
    return MSO_COLOR_TYPE.RGB;
  }

  CT_Color? get _color {
    final rPr = _r.rPr;
    return rPr?.color;
  }

  CT_R get _r {
    final base = element;
    if (base is CT_R) {
      return base;
    }
    throw StateError('ColorFormat requires a CT_R parent but got ${base.runtimeType}');
  }
}
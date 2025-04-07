/// Path: lib/src/oxml/text/font.dart
/// Based on python-docx: docx/oxml/text/font.py
/// Custom element classes related to run properties (font).

import 'package:xml/xml.dart';

import '../../enum/dml.dart' as dml; // Use prefix for MSO_THEME_COLOR
import '../../enum/text.dart'; // For WD_COLOR_INDEX, WD_UNDERLINE
import '../../shared.dart'; // For Length
import '../ns.dart';
import '../parser.dart';
import '../shared.dart'; // For CT_OnOff, CT_String
import '../simpletypes.dart';
import '../xmlchemy.dart';

/// `<w:color>` element, specifying the color of a font and perhaps other objects.
class CT_Color extends BaseOxmlElement {
  CT_Color(super.element);

  /// The color value, either an RGB hex string or "auto".
  /// Uses [Object] because `fromXml` can return `String` or `RGBColor`.
  Object get val => getReqAttrVal('w:val', stHexColorConverter);
  set val(Object value) => setReqAttrVal('w:val', value, stHexColorConverter);

  /// The theme color, if specified.
  dml.MSO_THEME_COLOR_INDEX? get themeColor =>
      getAttrVal('w:themeColor', msoThemeColorIndexConverter);
  set themeColor(dml.MSO_THEME_COLOR_INDEX? value) =>
      setAttrVal('w:themeColor', value, msoThemeColorIndexConverter);

  /// Creates a new `XmlElement` suitable for a CT_Color.
  static XmlElement create(String nsptagname, Object val) {
    final xmlVal = stHexColorConverter.toXml(val);
    if (xmlVal == null) {
      // This should ideally not happen for 'auto' or RGBColor, but check anyway
      throw ArgumentError("Cannot create CT_Color with null xml value from $val");
    }
    return OxmlElement(nsptagname, attrs: {qn('w:val'): xmlVal});
  }
}

/// `<w:rFonts>` element, specifying typeface names for various language types.
class CT_Fonts extends BaseOxmlElement {
  CT_Fonts(super.element);

  /// Font name for ASCII characters.
  String? get ascii => getAttrVal('w:ascii', stStringConverter);
  set ascii(String? value) => setAttrVal('w:ascii', value, stStringConverter);

  /// Font name for High ANSI characters.
  String? get hAnsi => getAttrVal('w:hAnsi', stStringConverter);
  set hAnsi(String? value) => setAttrVal('w:hAnsi', value, stStringConverter);

  // Other attributes like w:eastAsia, w:cs can be added here if needed.
}

/// `<w:highlight>` element, specifying font highlighting/background color.
class CT_Highlight extends BaseOxmlElement {
  CT_Highlight(super.element);

  /// The highlight color.
  WD_COLOR_INDEX get val => getReqAttrVal('w:val', wdColorIndexConverter);
  set val(WD_COLOR_INDEX value) =>
      setReqAttrVal('w:val', value, wdColorIndexConverter);

  /// Creates a new `XmlElement` suitable for a CT_Highlight.
  static XmlElement create(String nsptagname, WD_COLOR_INDEX val) {
    final xmlVal = wdColorIndexConverter.toXml(val);
    if (xmlVal == null) {
      throw ArgumentError(
          "Cannot create CT_Highlight with null xml value from $val");
    }
    return OxmlElement(nsptagname, attrs: {qn('w:val'): xmlVal});
  }
}

/// Used for `<w:sz>` element and others, specifying font size in half-points.
class CT_HpsMeasure extends BaseOxmlElement {
  CT_HpsMeasure(super.element);

  /// Font size as a [Length] object.
  Length get val => getReqAttrVal('w:val', stHpsMeasureConverter);
  set val(Length value) => setReqAttrVal('w:val', value, stHpsMeasureConverter);

  /// Creates a new `XmlElement` suitable for a CT_HpsMeasure.
  static XmlElement create(String nsptagname, Length val) {
    final xmlVal = stHpsMeasureConverter.toXml(val);
    if (xmlVal == null) {
      throw ArgumentError(
          "Cannot create CT_HpsMeasure with null xml value from $val");
    }
    return OxmlElement(nsptagname, attrs: {qn('w:val'): xmlVal});
  }
}

/// `<w:rPr>` element, containing the properties for a run.
class CT_RPr extends BaseOxmlElement {
  CT_RPr(super.element);

  // Define tag names for properties (match Python's _tag_seq for ordering hints)
  static const String _wRStyle = "w:rStyle";
  static const String _wRFonts = "w:rFonts";
  static const String _wB = "w:b";
  static const String _wBCs = "w:bCs";
  static const String _wI = "w:i";
  static const String _wICs = "w:iCs";
  static const String _wCaps = "w:caps";
  static const String _wSmallCaps = "w:smallCaps";
  static const String _wStrike = "w:strike";
  static const String _wDstrike = "w:dstrike";
  static const String _wOutline = "w:outline";
  static const String _wShadow = "w:shadow";
  static const String _wEmboss = "w:emboss";
  static const String _wImprint = "w:imprint";
  static const String _wNoProof = "w:noProof";
  static const String _wSnapToGrid = "w:snapToGrid";
  static const String _wVanish = "w:vanish";
  static const String _wWebHidden = "w:webHidden";
  static const String _wColor = "w:color";
  static const String _wSz = "w:sz";
  static const String _wHighlight = "w:highlight";
  static const String _wU = "w:u";
  static const String _wVertAlign = "w:vertAlign";
  static const String _wRtl = "w:rtl";
  static const String _wCs = "w:cs";
  static const String _wSpecVanish = "w:specVanish";
  static const String _wOMath = "w:oMath";

  // Define successor lists based on Python's _tag_seq
  // Use 'final' instead of 'const' for lists using methods like sublist
  static final List<String> _rStyleSuccessors = [
    _wRFonts, _wB, _wBCs, _wI, _wICs, _wCaps, _wSmallCaps, _wStrike, _wDstrike,
    _wOutline, _wShadow, _wEmboss, _wImprint, _wNoProof, _wSnapToGrid, _wVanish,
    _wWebHidden, _wColor, _wSz, _wHighlight, _wU, _wVertAlign, _wRtl, _wCs,
    _wSpecVanish, _wOMath,
  ];
  static final List<String> _rFontsSuccessors = _rStyleSuccessors.sublist(1);
  static final List<String> _bSuccessors = _rFontsSuccessors.sublist(1);
  static final List<String> _bCsSuccessors = _bSuccessors.sublist(1);
  static final List<String> _iSuccessors = _bCsSuccessors.sublist(1);
  static final List<String> _iCsSuccessors = _iSuccessors.sublist(1);
  static final List<String> _capsSuccessors = _iCsSuccessors.sublist(1);
  static final List<String> _smallCapsSuccessors = _capsSuccessors.sublist(1);
  static final List<String> _strikeSuccessors = _smallCapsSuccessors.sublist(1);
  static final List<String> _dstrikeSuccessors = _strikeSuccessors.sublist(1);
  static final List<String> _outlineSuccessors = _dstrikeSuccessors.sublist(1);
  static final List<String> _shadowSuccessors = _outlineSuccessors.sublist(1);
  static final List<String> _embossSuccessors = _shadowSuccessors.sublist(1);
  static final List<String> _imprintSuccessors = _embossSuccessors.sublist(1);
  static final List<String> _noProofSuccessors = _imprintSuccessors.sublist(1);
  static final List<String> _snapToGridSuccessors = _noProofSuccessors.sublist(1);
  static final List<String> _vanishSuccessors = _snapToGridSuccessors.sublist(1);
  static final List<String> _webHiddenSuccessors = _vanishSuccessors.sublist(1);
  static final List<String> _colorSuccessors = _webHiddenSuccessors.sublist(1);
  static final List<String> _szSuccessors = _colorSuccessors.sublist(1);
  static final List<String> _highlightSuccessors = _szSuccessors.sublist(1);
  static final List<String> _uSuccessors = _highlightSuccessors.sublist(1);
  static final List<String> _vertAlignSuccessors = _uSuccessors.sublist(1);
  static final List<String> _rtlSuccessors = _vertAlignSuccessors.sublist(1);
  static final List<String> _csSuccessors = _rtlSuccessors.sublist(1);
  static final List<String> _specVanishSuccessors = _csSuccessors.sublist(1);
  static final List<String> _oMathSuccessors = _specVanishSuccessors.sublist(1);


  // --- Getters for child elements (using helpers below) ---
  CT_String? get rStyle => _getCTString(_wRStyle);
  CT_Fonts? get rFonts => _getCTFonts(_wRFonts);
  CT_Color? get color => _getCTColor(_wColor);
  CT_HpsMeasure? get sz => _getCTHpsMeasure(_wSz);
  CT_Highlight? get highlight => _getCTHighlight(_wHighlight);
  CT_Underline? get u => _getCTUnderline(_wU);
  CT_VerticalAlignRun? get vertAlign => _getCTVerticalAlignRun(_wVertAlign);
  // Boolean getters are handled by explicit properties further down


  // --- Helper getters for wrapping child elements ---
  CT_String? _getCTString(String tag) =>
      childOrNull(tag) == null ? null : CT_String(childOrNull(tag)!);
  CT_Fonts? _getCTFonts(String tag) =>
      childOrNull(tag) == null ? null : CT_Fonts(childOrNull(tag)!);
  CT_OnOff? _getCTOnOff(String tag) =>
      childOrNull(tag) == null ? null : CT_OnOff(childOrNull(tag)!);
  CT_Color? _getCTColor(String tag) =>
      childOrNull(tag) == null ? null : CT_Color(childOrNull(tag)!);
  CT_HpsMeasure? _getCTHpsMeasure(String tag) =>
      childOrNull(tag) == null ? null : CT_HpsMeasure(childOrNull(tag)!);
  CT_Highlight? _getCTHighlight(String tag) =>
      childOrNull(tag) == null ? null : CT_Highlight(childOrNull(tag)!);
  CT_Underline? _getCTUnderline(String tag) =>
      childOrNull(tag) == null ? null : CT_Underline(childOrNull(tag)!);
  CT_VerticalAlignRun? _getCTVerticalAlignRun(String tag) =>
      childOrNull(tag) == null
          ? null
          : CT_VerticalAlignRun(childOrNull(tag)!);


  // --- Get or Add methods ---
  CT_Highlight getOrAddHighlight() => CT_Highlight(
      getOrAddChild(_wHighlight, _highlightSuccessors, _factoryHighlight));
  CT_Fonts getOrAddRFonts() =>
      CT_Fonts(getOrAddChild(_wRFonts, _rFontsSuccessors, _factoryRFonts));
  CT_HpsMeasure getOrAddSz() =>
      CT_HpsMeasure(getOrAddChild(_wSz, _szSuccessors, _factorySz));
  CT_VerticalAlignRun getOrAddVertAlign() => CT_VerticalAlignRun(
      getOrAddChild(_wVertAlign, _vertAlignSuccessors, _factoryVertAlign));
  // CT_String getOrAddRStyle() => CT_String(
  //    getOrAddChild(_wRStyle, _rStyleSuccessors, _factoryRStyle)); // Renamed from addRStyle for consistency
  CT_Underline getOrAddU() => CT_Underline(
      getOrAddChild(_wU, _uSuccessors, _factoryU)); // Renamed from addU for consistency


  // --- Remover methods ---
  void removeHighlight() => removeChild(_wHighlight);
  void removeRFonts() => removeChild(_wRFonts);
  void removeRStyle() => removeChild(_wRStyle);
  void removeSz() => removeChild(_wSz);
  void removeU() => removeChild(_wU);
  void removeVertAlign() => removeChild(_wVertAlign);


  // --- Private Element Factories ---
  static XmlElement _factoryRStyle() => CT_String.create(_wRStyle, ''); // Needs initial val
  static XmlElement _factoryRFonts() => OxmlElement(_wRFonts);
  static XmlElement _factorySz() => CT_HpsMeasure.create(_wSz, Length(0)); // Now calls class method
  static XmlElement _factoryHighlight() => CT_Highlight.create(_wHighlight, WD_COLOR_INDEX.AUTO); // Now calls class method
  static XmlElement _factoryU() => CT_Underline.create(_wU, WD_UNDERLINE.INHERITED); // Now calls class method
  static XmlElement _factoryVertAlign() => CT_VerticalAlignRun.create(_wVertAlign, ST_VerticalAlignRunConverter.baseline); // Now calls class method
  static XmlElement factoryColor() {
    return CT_Color.create(_wColor, '000000'); // Now calls class method
  }
  static XmlElement _factoryOnOff(String tag) => OxmlElement(tag); // General factory


  // --- Property implementations using helpers ---

  WD_COLOR_INDEX? get highlightVal => highlight?.val;
  set highlightVal(WD_COLOR_INDEX? value) {
    if (value == null) {
      removeHighlight();
    } else {
      getOrAddHighlight().val = value;
    }
  }

  String? get rFontsAscii => rFonts?.ascii;
  set rFontsAscii(String? value) {
    if (value == null) {
      if (rFonts?.hAnsi == null) {
        removeRFonts();
      } else {
        rFonts?.ascii = null;
      }
    } else {
      getOrAddRFonts().ascii = value;
      // Also set hAnsi if it's null or currently same as ascii
      final currentHAnsi = rFonts?.hAnsi; // Read only once
      if (currentHAnsi == null || currentHAnsi == rFonts?.ascii) {
        getOrAddRFonts().hAnsi = value;
      }
    }
  }

  String? get rFontsHAnsi => rFonts?.hAnsi;
  set rFontsHAnsi(String? value) {
     if (value == null) {
        if (rFonts?.ascii == null) {
           removeRFonts();
        } else {
           rFonts?.hAnsi = null;
        }
     } else {
        getOrAddRFonts().hAnsi = value;
     }
  }

  String? get style => rStyle?.val;
  set style(String? value) {
    if (value == null) {
      removeRStyle();
    } else {
      // Use getOrAddChild directly for CT_String case
      final rStyleElement = getOrAddChild(_wRStyle, _rStyleSuccessors, _factoryRStyle);
      CT_String(rStyleElement).val = value;
    }
  }

  bool? get subscript {
    final va = vertAlign;
    if (va == null) return null;
    return va.val == ST_VerticalAlignRunConverter.subscript;
  }

  set subscript(bool? value) {
    if (value == null) {
      if (vertAlign?.val == ST_VerticalAlignRunConverter.subscript) {
        removeVertAlign();
      }
    } else if (value) {
      getOrAddVertAlign().val = ST_VerticalAlignRunConverter.subscript;
    } else {
      if (vertAlign?.val == ST_VerticalAlignRunConverter.subscript) {
        removeVertAlign();
      }
    }
  }

  bool? get superscript {
    final va = vertAlign;
    if (va == null) return null;
    return va.val == ST_VerticalAlignRunConverter.superscript;
  }

  set superscript(bool? value) {
    if (value == null) {
       if (vertAlign?.val == ST_VerticalAlignRunConverter.superscript) {
         removeVertAlign();
       }
    } else if (value) {
      getOrAddVertAlign().val = ST_VerticalAlignRunConverter.superscript;
    } else {
      if (vertAlign?.val == ST_VerticalAlignRunConverter.superscript) {
        removeVertAlign();
      }
    }
  }


  Length? get szVal => sz?.val;
  set szVal(Length? value) {
    if (value == null) {
      removeSz();
    } else {
      getOrAddSz().val = value;
    }
  }

  WD_UNDERLINE? get uVal => u?.val;
  set uVal(WD_UNDERLINE? value) {
    // Treat null and INHERITED the same (remove element)
    if (value == null || value == WD_UNDERLINE.INHERITED) {
       removeU();
    } else {
       // Use getOrAddU to ensure element exists, then set value
       getOrAddU().val = value;
    }
  }

  // --- Boolean property helper (internal) ---
  CT_OnOff? _getOnOffElement(String tagName) => _getCTOnOff(tagName);
  bool? _getBoolProp(String tagName) => _getOnOffElement(tagName)?.val;
  void _setBoolProp(String tagName, bool? value, List<String> successors) {
     if (value == null) {
        removeChild(tagName);
     } else {
        final onOffElement = getOrAddChild(tagName, successors, () => _factoryOnOff(tagName));
        CT_OnOff(onOffElement).val = value;
     }
  }

  // --- Explicit boolean properties using the helper ---
  bool? get bold => _getBoolProp(_wB);
  set bold(bool? value) => _setBoolProp(_wB, value, _bSuccessors);

  bool? get italic => _getBoolProp(_wI);
  set italic(bool? value) => _setBoolProp(_wI, value, _iSuccessors);

  bool? get caps => _getBoolProp(_wCaps);
  set caps(bool? value) => _setBoolProp(_wCaps, value, _capsSuccessors);

  bool? get smallCaps => _getBoolProp(_wSmallCaps);
  set smallCaps(bool? value) => _setBoolProp(_wSmallCaps, value, _smallCapsSuccessors);

  bool? get strike => _getBoolProp(_wStrike);
  set strike(bool? value) => _setBoolProp(_wStrike, value, _strikeSuccessors);

  bool? get doubleStrike => _getBoolProp(_wDstrike);
  set doubleStrike(bool? value) => _setBoolProp(_wDstrike, value, _dstrikeSuccessors);

  bool? get outline => _getBoolProp(_wOutline);
  set outline(bool? value) => _setBoolProp(_wOutline, value, _outlineSuccessors);

  bool? get shadow => _getBoolProp(_wShadow);
  set shadow(bool? value) => _setBoolProp(_wShadow, value, _shadowSuccessors);

  bool? get emboss => _getBoolProp(_wEmboss);
  set emboss(bool? value) => _setBoolProp(_wEmboss, value, _embossSuccessors);

  bool? get imprint => _getBoolProp(_wImprint);
  set imprint(bool? value) => _setBoolProp(_wImprint, value, _imprintSuccessors);

  bool? get noProof => _getBoolProp(_wNoProof);
  set noProof(bool? value) => _setBoolProp(_wNoProof, value, _noProofSuccessors);

  bool? get snapToGrid => _getBoolProp(_wSnapToGrid);
  set snapToGrid(bool? value) => _setBoolProp(_wSnapToGrid, value, _snapToGridSuccessors);

  bool? get vanish => _getBoolProp(_wVanish); // aka hidden
  set vanish(bool? value) => _setBoolProp(_wVanish, value, _vanishSuccessors);

  bool? get webHidden => _getBoolProp(_wWebHidden);
  set webHidden(bool? value) => _setBoolProp(_wWebHidden, value, _webHiddenSuccessors);

  bool? get rtl => _getBoolProp(_wRtl);
  set rtl(bool? value) => _setBoolProp(_wRtl, value, _rtlSuccessors);

  bool? get complexScript => _getBoolProp(_wCs); // aka cs
  set complexScript(bool? value) => _setBoolProp(_wCs, value, _csSuccessors);

  bool? get specVanish => _getBoolProp(_wSpecVanish);
  set specVanish(bool? value) => _setBoolProp(_wSpecVanish, value, _specVanishSuccessors);

  bool? get oMath => _getBoolProp(_wOMath);
  set oMath(bool? value) => _setBoolProp(_wOMath, value, _oMathSuccessors);

  // cs_bold (bCs) and cs_italic (iCs) would follow the same pattern if needed
  bool? get csBold => _getBoolProp(_wBCs);
  set csBold(bool? value) => _setBoolProp(_wBCs, value, _bCsSuccessors);
  bool? get csItalic => _getBoolProp(_wICs);
  set csItalic(bool? value) => _setBoolProp(_wICs, value, _iCsSuccessors);

}

/// `<w:u>` element, specifying the underlining style for a run.
class CT_Underline extends BaseOxmlElement {
  CT_Underline(super.element);

  WD_UNDERLINE? get val => getAttrVal('w:val', wdUnderlineConverter, defaultValue: WD_UNDERLINE.INHERITED);
  // Use INHERITED as default, setting null removes attribute
  set val(WD_UNDERLINE? value) => setAttrVal('w:val', value, wdUnderlineConverter, defaultValue: WD_UNDERLINE.INHERITED);

   /// Creates a new `XmlElement` suitable for a CT_Underline.
   static XmlElement create(String nsptagname, WD_UNDERLINE? val) { // Accept null
      final attrs = <String, String>{};
      // Only add the attribute if the value is not null and not the default INHERITED
      if (val != null && val != WD_UNDERLINE.INHERITED) {
         final xmlVal = wdUnderlineConverter.toXml(val);
         if (xmlVal != null) {
             attrs[qn('w:val')] = xmlVal;
         }
      }
      return OxmlElement(nsptagname, attrs: attrs);
   }
}

/// `<w:vertAlign>` element, specifying subscript or superscript.
class CT_VerticalAlignRun extends BaseOxmlElement {
  CT_VerticalAlignRun(super.element);

  String get val => getReqAttrVal('w:val', stVerticalAlignRunConverter);
  set val(String value) => setReqAttrVal('w:val', value, stVerticalAlignRunConverter);

  /// Creates a new `XmlElement` suitable for a CT_VerticalAlignRun.
  static XmlElement create(String nsptagname, String val) {
      // Basic validation could be added here if desired
      stVerticalAlignRunConverter.validate(val);
      return OxmlElement(nsptagname, attrs: {qn('w:val'): val});
  }
}
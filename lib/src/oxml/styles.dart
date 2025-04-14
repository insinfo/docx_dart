/// Path: lib/src/oxml/styles.dart
/// Based on python-docx: docx/oxml/styles.py
/// Custom element classes related to the styles part (<w:styles>, <w:style>, etc.).

import 'package:docx_dart/src/oxml/shared.dart';
import 'package:xml/xml.dart';
import '../enum/style.dart'; // WD_STYLE_TYPE
import 'ns.dart' show qn; // qn function
import 'parser.dart' show OxmlElement; // OxmlElement factory
import 'simpletypes.dart';
import 'text/parfmt.dart' show CT_PPr;
import 'text/font.dart' show CT_RPr;
import 'xmlchemy.dart' show BaseOxmlElement; // BaseOxmlElement

// --- styleIdFromName helper ---
String styleIdFromName(String name) {
  const styleAliases = {
    "Caption": "Caption",
    "Heading 1": "Heading1",
    "Heading 2": "Heading2",
    "Heading 3": "Heading3",
    "Heading 4": "Heading4",
    "Heading 5": "Heading5",
    "Heading 6": "Heading6",
    "Heading 7": "Heading7",
    "Heading 8": "Heading8",
    "Heading 9": "Heading9",
  };
  return styleAliases[name] ?? name.replaceAll(' ', '');
}

/// `<w:latentStyles>` element ...
class CT_LatentStyles extends BaseOxmlElement {
  CT_LatentStyles(super.element);

  /// Creates a new `<w:latentStyles>` element.
  static XmlElement create() => OxmlElement(qnTagName);

  // ... (Properties: lsdExceptionElements, count, def*, boolProp, setBoolProp remain the same) ...
  List<CT_LsdException> get lsdExceptionElements =>
      childrenWhereType<CT_LsdException>(
          CT_LsdException.qnTagName, (el) => CT_LsdException(el));
  List<CT_LsdException> get lsdException_lst => lsdExceptionElements;
  int? get count => getAttrVal('w:count', stDecimalNumberConverter);
  set count(int? value) =>
      setAttrVal('w:count', value, stDecimalNumberConverter);
  bool? get defLockedState => getAttrVal('w:defLockedState', stOnOffConverter);
  set defLockedState(bool? value) =>
      setAttrVal('w:defLockedState', value, stOnOffConverter);
  bool? get defQFormat => getAttrVal('w:defQFormat', stOnOffConverter);
  set defQFormat(bool? value) =>
      setAttrVal('w:defQFormat', value, stOnOffConverter);
  bool? get defSemiHidden => getAttrVal('w:defSemiHidden', stOnOffConverter);
  set defSemiHidden(bool? value) =>
      setAttrVal('w:defSemiHidden', value, stOnOffConverter);
  int? get defUIPriority =>
      getAttrVal('w:defUIPriority', stDecimalNumberConverter);
  set defUIPriority(int? value) =>
      setAttrVal('w:defUIPriority', value, stDecimalNumberConverter);
  bool? get defUnhideWhenUsed =>
      getAttrVal('w:defUnhideWhenUsed', stOnOffConverter);
  set defUnhideWhenUsed(bool? value) =>
      setAttrVal('w:defUnhideWhenUsed', value, stOnOffConverter);
  bool boolProp(String attrName) =>
      getAttrVal(attrName, stOnOffConverter) ?? false;
  void setBoolProp(String attrName, bool? value) =>
      setAttrVal(attrName, value, stOnOffConverter, defaultValue: false);

  /// Return the `<w:lsdException>` child element matching [internalName]. Null if not found.
  CT_LsdException? getByName(String internalName) {
    final qName = CT_LsdException.qnTagName;
    for (final child in childrenWhereType<CT_LsdException>(
        qName, (e) => CT_LsdException(e))) {
      if (child.name == internalName) return child;
    }
    return null;
  }

  /// Add a new `<w:lsdException>` child element.
  CT_LsdException addLsdException() {
    // --- CORRECTION: Call create without name, set name later if needed ---
    final lsdExceptionElement = CT_LsdException.create();
    // --- End Correction ---
    addChildElement([], () => lsdExceptionElement);
    return CT_LsdException(lsdExceptionElement);
  }

  CT_LsdException add_lsdException() => addLsdException();

  static final qnTagName = qn('w:latentStyles');
}

/// `<w:lsdException>` element ...
class CT_LsdException extends BaseOxmlElement {
  CT_LsdException(super.element);

  /// Creates a new `<w:lsdException>` element. Name is required by schema but
  /// can be set after creation via the setter.
  // --- CORRECTION: Make name optional here, rely on setter ---
  static XmlElement create({String? name}) {
    final attrs = <String, String>{};
    if (name != null) {
      attrs[qn('w:name')] = name;
    }
    return OxmlElement(qnTagName, attrs: attrs);
  }
  // --- End Correction ---

  // ... (Properties: locked, name, qFormat, semiHidden, uiPriority, unhideWhenUsed remain the same) ...
  bool? get locked => getAttrVal('w:locked', stOnOffConverter);
  set locked(bool? value) => setAttrVal('w:locked', value, stOnOffConverter);
  String get name => getReqAttrVal('w:name', stStringConverter);
  set name(String value) => setReqAttrVal('w:name', value, stStringConverter);
  bool? get qFormat => getAttrVal('w:qFormat', stOnOffConverter);
  set qFormat(bool? value) => setAttrVal('w:qFormat', value, stOnOffConverter);
  bool? get semiHidden => getAttrVal('w:semiHidden', stOnOffConverter);
  set semiHidden(bool? value) =>
      setAttrVal('w:semiHidden', value, stOnOffConverter);
  int? get uiPriority => getAttrVal('w:uiPriority', stDecimalNumberConverter);
  set uiPriority(int? value) =>
      setAttrVal('w:uiPriority', value, stDecimalNumberConverter);
  bool? get unhideWhenUsed => getAttrVal('w:unhideWhenUsed', stOnOffConverter);
  set unhideWhenUsed(bool? value) =>
      setAttrVal('w:unhideWhenUsed', value, stOnOffConverter);

  void delete() {
    final parent = element.parent;
    if (parent != null) {
      parent.children.remove(element);
    }
  }

  bool? onOffProp(String attrName) => getAttrVal(attrName, stOnOffConverter);
  void setOnOffProp(String attrName, bool? value) =>
      setAttrVal(attrName, value, stOnOffConverter);

  static final qnTagName = qn('w:lsdException');
}

/// `<w:style>` element ...
class CT_Style extends BaseOxmlElement {
  CT_Style(super.element);

  /// Creates a new `<w:style>` element. Requires type and styleId.
  static XmlElement create(
          {required WD_STYLE_TYPE type, required String styleId}) =>
      OxmlElement(qnTagName, attrs: {
        qn('w:type'): wdStyleTypeConverter.toXml(type)!,
        qn('w:styleId'): styleId,
      });

  // ... (_tagSeq, Attribute Getters/Setters remain the same) ...
  static final List<String> _tagSeq = [
    qn("w:name"),
    qn("w:aliases"),
    qn("w:basedOn"),
    qn("w:next"),
    qn("w:link"),
    qn("w:autoRedefine"),
    qn("w:hidden"),
    qn("w:uiPriority"),
    qn("w:semiHidden"),
    qn("w:unhideWhenUsed"),
    qn("w:qFormat"),
    qn("w:locked"),
    qn("w:personal"),
    qn("w:personalCompose"),
    qn("w:personalReply"),
    qn("w:rsid"),
    qn("w:pPr"),
    qn("w:rPr"),
    qn("w:tblPr"),
    qn("w:trPr"),
    qn("w:tcPr"),
    qn("w:tblStylePr"),
  ];
  WD_STYLE_TYPE? get type => getAttrVal('w:type', wdStyleTypeConverter);
  set type(WD_STYLE_TYPE? value) =>
      setAttrVal('w:type', value, wdStyleTypeConverter);
  String? get styleId => getAttrVal('w:styleId', stStringConverter);
  set styleId(String? value) =>
      setAttrVal('w:styleId', value, stStringConverter);
  bool? get defaultAttr => getAttrVal('w:default', stOnOffConverter);
  set defaultAttr(bool? value) =>
      setAttrVal('w:default', value, stOnOffConverter);
  bool? get customStyle => getAttrVal('w:customStyle', stOnOffConverter);
  set customStyle(bool? value) =>
      setAttrVal('w:customStyle', value, stOnOffConverter);

  // ... (Child Element Getters remain the same) ...
  CT_String? get nameElement => _getCTString(_tagSeq[0]);
  CT_String? get basedOnElement => _getCTString(_tagSeq[2]);
  CT_String? get nextElement => _getCTString(_tagSeq[3]);
  CT_DecimalNumber? get uiPriorityElement => _getCTDecimalNumber(_tagSeq[7]);
  CT_OnOff? get semiHiddenElement => _getCTOnOff(_tagSeq[8]);
  CT_OnOff? get unhideWhenUsedElement => _getCTOnOff(_tagSeq[9]);
  CT_OnOff? get qFormatElement => _getCTOnOff(_tagSeq[10]);
  CT_OnOff? get lockedElement => _getCTOnOff(_tagSeq[11]);
  CT_PPr? get pPrElement => _getCTPPr(_tagSeq[16]);
  CT_RPr? get rPrElement => _getCTRPr(_tagSeq[17]);

  // ... (Helper Getters remain the same) ...
  CT_String? _getCTString(String tag) =>
      childOrNull(tag) == null ? null : CT_String(childOrNull(tag)!);
  CT_DecimalNumber? _getCTDecimalNumber(String tag) =>
      childOrNull(tag) == null ? null : CT_DecimalNumber(childOrNull(tag)!);
  CT_OnOff? _getCTOnOff(String tag) =>
      childOrNull(tag) == null ? null : CT_OnOff(childOrNull(tag)!);
  CT_PPr? _getCTPPr(String tag) =>
      childOrNull(tag) == null ? null : CT_PPr(childOrNull(tag)!);
  CT_RPr? _getCTRPr(String tag) =>
      childOrNull(tag) == null ? null : CT_RPr(childOrNull(tag)!);

  // ... (Get or Add Methods remain the same) ...
  CT_String getOrAddName() => CT_String(getOrAddChild(
      _tagSeq[0], _tagSeq.sublist(1), () => CT_String.create(_tagSeq[0], '')));
  CT_String getOrAddBasedOn() => CT_String(getOrAddChild(
      _tagSeq[2], _tagSeq.sublist(3), () => CT_String.create(_tagSeq[2], '')));
  CT_String getOrAddNext() => CT_String(getOrAddChild(
      _tagSeq[3], _tagSeq.sublist(4), () => CT_String.create(_tagSeq[3], '')));
  CT_DecimalNumber getOrAddUiPriority() => CT_DecimalNumber(getOrAddChild(
      _tagSeq[7],
      _tagSeq.sublist(8),
      () => CT_DecimalNumber.create(_tagSeq[7], 0)));
  CT_OnOff getOrAddSemiHidden() => CT_OnOff(getOrAddChild(
      _tagSeq[8], _tagSeq.sublist(9), () => CT_OnOff.create(_tagSeq[8])));
  CT_OnOff getOrAddUnhideWhenUsed() => CT_OnOff(getOrAddChild(
      _tagSeq[9], _tagSeq.sublist(10), () => CT_OnOff.create(_tagSeq[9])));
  CT_OnOff getOrAddQFormat() => CT_OnOff(getOrAddChild(
      _tagSeq[10], _tagSeq.sublist(11), () => CT_OnOff.create(_tagSeq[10])));
  CT_OnOff getOrAddLocked() => CT_OnOff(getOrAddChild(
      _tagSeq[11], _tagSeq.sublist(12), () => CT_OnOff.create(_tagSeq[11])));
  CT_PPr getOrAddPPr() =>
      CT_PPr(getOrAddChild(_tagSeq[16], _tagSeq.sublist(17), CT_PPr.create));
  CT_RPr getOrAddRPr() =>
      CT_RPr(getOrAddChild(_tagSeq[17], _tagSeq.sublist(18), CT_RPr.create));

  // ... (Remove Methods remain the same) ...
  void removeName() => removeChild(_tagSeq[0]);
  void removeBasedOn() => removeChild(_tagSeq[2]);
  void removeNext() => removeChild(_tagSeq[3]);
  void removeUiPriority() => removeChild(_tagSeq[7]);
  void removeSemiHidden() => removeChild(_tagSeq[8]);
  void removeUnhideWhenUsed() => removeChild(_tagSeq[9]);
  void removeQFormat() => removeChild(_tagSeq[10]);
  void removeLocked() => removeChild(_tagSeq[11]);
  void removePPr() => removeChild(_tagSeq[16]);
  void removeRPr() => removeChild(_tagSeq[17]);

  // --- Higher-Level Properties ---

  String? get basedOnVal => basedOnElement?.val;
  set basedOnVal(String? value) {
    if (value == null)
      removeBasedOn();
    else
      getOrAddBasedOn().val = value;
  }

  CT_Style? get baseStyle {
    final basedOnId = basedOnVal;
    if (basedOnId == null) return null;
    final stylesNode = element.parent; // Returns XmlNode?

    // --- CORRECTION: Check parent type and qualified name ---
    if (stylesNode is! XmlElement ||
        stylesNode.name.qualified != CT_Styles.qnTagName) {
      return null; // Parent must be the <w:styles> element
    }
    // --- End Correction ---

    // --- CORRECTION: Pass XmlElement to CT_Styles ---
    return CT_Styles(stylesNode).getById(basedOnId);
    // --- End Correction ---
  }

  bool get builtin => customStyle != true;

  void delete() {
    final parent = element.parent;
    if (parent != null) parent.children.remove(element);
  }

  bool get lockedVal => lockedElement?.val ?? false;
  set lockedVal(bool value) {
    if (!value)
      removeLocked();
    else
      getOrAddLocked().val = true;
  }

  String? get nameVal => nameElement?.val;
  set nameVal(String? value) {
    if (value == null)
      removeName();
    else
      getOrAddName().val = value;
  }

  CT_Style? get nextStyle {
    final nextStyleId = nextElement?.val;
    if (nextStyleId == null) return null;
    final stylesNode = element.parent; // Returns XmlNode?

    // --- CORRECTION: Check parent type and qualified name ---
    if (stylesNode is! XmlElement ||
        stylesNode.name.qualified != CT_Styles.qnTagName) {
      return null;
    }
    // --- End Correction ---

    // --- CORRECTION: Pass XmlElement to CT_Styles ---
    return CT_Styles(stylesNode).getById(nextStyleId);
    // --- End Correction ---
  }

  bool get qFormatVal => qFormatElement?.val ?? false;
  set qFormatVal(bool value) {
    if (!value)
      removeQFormat();
    else
      getOrAddQFormat().val = true;
  }

  bool get semiHiddenVal => semiHiddenElement?.val ?? false;
  set semiHiddenVal(bool value) {
    if (!value)
      removeSemiHidden();
    else
      getOrAddSemiHidden().val = true;
  }

  int? get uiPriorityVal => uiPriorityElement?.val;
  set uiPriorityVal(int? value) {
    if (value == null)
      removeUiPriority();
    else
      getOrAddUiPriority().val = value;
  }

  bool get unhideWhenUsedVal => unhideWhenUsedElement?.val ?? false;
  set unhideWhenUsedVal(bool value) {
    if (!value)
      removeUnhideWhenUsed();
    else
      getOrAddUnhideWhenUsed().val = true;
  }

  static final qnTagName = qn('w:style');
}

/// `<w:styles>` element ...
class CT_Styles extends BaseOxmlElement {
  CT_Styles(super.element);

  /// Creates a new `<w:styles>` element.
  static XmlElement create() => OxmlElement(qnTagName);

  // ... (Getters for latentStylesElement, styleElements remain the same) ...
  CT_LatentStyles? get latentStylesElement =>
      childOrNull(CT_LatentStyles.qnTagName) != null
          ? CT_LatentStyles(childOrNull(CT_LatentStyles.qnTagName)!)
          : null;
  List<CT_Style> get styleElements =>
      childrenWhereType<CT_Style>(CT_Style.qnTagName, (el) => CT_Style(el));
  List<CT_Style> get style_lst => styleElements;

  CT_LatentStyles getOrAddLatentStyles() => CT_LatentStyles(getOrAddChild(
      CT_LatentStyles.qnTagName, [CT_Style.qnTagName], CT_LatentStyles.create));

  CT_Style addStyleOfType(String name, WD_STYLE_TYPE styleType,
      {bool builtin = false}) {
    final internalName = name;
    final styleId = styleIdFromName(name);

    final styleElement = CT_Style.create(type: styleType, styleId: styleId);
    final newStyle = CT_Style(styleElement);

    newStyle.nameVal = internalName;
    newStyle.customStyle = builtin ? null : true;

    addChildElement([], () => styleElement);

    return newStyle;
  }

  CT_Style? defaultFor(WD_STYLE_TYPE styleType) {
    for (final style in styleElements.reversed) {
      if (style.type == styleType && style.defaultAttr == true) {
        return style;
      }
    }
    return null;
  }

  CT_Style? getById(String styleId) {
    for (final style in styleElements) {
      if (style.styleId == styleId) {
        return style;
      }
    }
    return null;
  }

  CT_Style? getByName(String internalName) {
    for (final style in styleElements) {
      if (style.nameVal == internalName) {
        return style;
      }
    }
    return null;
  }

  static final qnTagName = qn('w:styles');
}

// --- Placeholder Converters ---
// ... (Definitions remain the same) ...
final stOnOffConverter = const ST_OnOffConverter();
final stDecimalNumberConverter = const ST_DecimalNumberConverter();
final stStringConverter = const ST_StringConverter();
final wdStyleTypeConverter = const WD_STYLE_TYPE_Converter();

class ST_OnOffConverter implements BaseSimpleType<bool> {
  const ST_OnOffConverter();
  @override
  bool fromXml(String xmlValue) =>
      ['1', 'true', 'on'].contains(xmlValue.toLowerCase());
  @override
  String? toXml(bool? value) => value == null ? null : (value ? '1' : '0');
  @override
  void validate(bool value) {}
}

class ST_DecimalNumberConverter implements BaseSimpleType<int> {
  const ST_DecimalNumberConverter();
  @override
  int fromXml(String xmlValue) => int.parse(xmlValue);
  @override
  String? toXml(int? value) => value?.toString();
  @override
  void validate(int value) {} // Add range checks if needed
}

class ST_StringConverter implements BaseSimpleType<String> {
  const ST_StringConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {}
}

class WD_STYLE_TYPE_Converter implements BaseSimpleType<WD_STYLE_TYPE> {
  const WD_STYLE_TYPE_Converter();
  @override
  WD_STYLE_TYPE fromXml(String xmlValue) =>
      WD_STYLE_TYPE.values.firstWhere((e) => e.xmlValue == xmlValue,
          orElse: () =>
              throw FormatException("Invalid WD_STYLE_TYPE value: $xmlValue"));
  @override
  String? toXml(WD_STYLE_TYPE? value) => value?.xmlValue;
  @override
  void validate(WD_STYLE_TYPE value) {}
}

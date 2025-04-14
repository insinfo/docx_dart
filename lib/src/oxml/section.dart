/// Path: lib/src/oxml/section.dart
/// Based on python-docx: docx/oxml/section.py
/// Custom element classes related to sections (<w:sectPr>, etc.).

import 'package:docx_dart/src/oxml/shared.dart';
import 'package:xml/xml.dart';
import 'package:collection/collection.dart'; // For firstWhereOrNull

import '../enum/section.dart'; // WD_HEADER_FOOTER, WD_ORIENTATION, WD_SECTION_START
import '../shared.dart' show Length, Twips; // Length, Emu, Twips

import 'simpletypes.dart';
import 'ns.dart';
import 'parser.dart';
import 'table.dart' show CT_Tbl;
import 'text/paragraph.dart' show CT_P;
import 'xmlchemy.dart' show BaseOxmlElement;
import 'xmlchemy_descriptors.dart'; // Descriptors like ZeroOrOne

// --- Instanciando os conversores (ASSUME que as classes estão definidas em simpletypes.dart) ---
final stSignedTwipsMeasureConverter = const ST_SignedTwipsMeasureConverter();
final stTwipsMeasureConverter = const ST_TwipsMeasureConverter();
final wdHeaderFooterIndexConverter = const WD_HEADER_FOOTER_Converter();
final xsdStringConverter = const XsdStringConverter();
final wdOrientationConverter = const WD_ORIENTATION_Converter();
final wdSectionStartConverter = const WD_SECTION_START_Converter();
// --- Fim dos Conversores ---

/// `<w:hdr>` or `<w:ftr>` element, root element for header/footer part.
class CT_HdrFtr extends BaseOxmlElement {
  CT_HdrFtr(super.element);

  /// Creates a new header/footer element with a default empty paragraph.
  static XmlElement create(String qnTagName) {
    // --- CORREÇÃO: Remover children, adicionar depois ---
    final hdrFtr = OxmlElement(qnTagName);
    hdrFtr.children.add(CT_P.create()); // Add required <w:p>
    return hdrFtr;
    // --- Fim da Correção ---
  }

  static final qnHdr = qn('w:hdr');
  static final qnFtr = qn('w:ftr');

  static final _p = ZeroOrMore<CT_P>(qn('w:p'), (el) => CT_P(el));
  static final _tbl = ZeroOrMore<CT_Tbl>(qn('w:tbl'), (el) => CT_Tbl(el));

  List<CT_P> get pList => _p.getElements(this);
  List<CT_P> get p_lst => pList;

  List<CT_Tbl> get tblList => _tbl.getElements(this);
  List<CT_Tbl> get tbl_lst => tblList;

  List<BaseOxmlElement> get innerContentElements {
    final content = <BaseOxmlElement>[];
    final pTag = CT_P.qnTagName;
    final tblTag = CT_Tbl.qnTagName;
    for (final child in element.children.whereType<XmlElement>()) {
      if (child.name.qualified == pTag) {
        content.add(CT_P(child));
      } else if (child.name.qualified == tblTag) {
        content.add(CT_Tbl(child));
      }
    }
    return content;
  }

  CT_P addP() {
    final pElement = CT_P.create();
    element.children.add(pElement);
    return CT_P(pElement);
  }

  CT_P add_p() => addP();

  void insertTbl(CT_Tbl tbl) {
    // Adiciona o elemento XML real, não o wrapper Dart
    if (tbl.element.parent != null) {
      tbl.element.parent!.children
          .remove(tbl.element); // Desanexa se já tiver pai
    }
    element.children.add(tbl.element);
  }

  void insert_tbl(CT_Tbl tbl) => insertTbl(tbl);
}

/// `<w:headerReference>` or `<w:footerReference>` element.
class CT_HdrFtrRef extends BaseOxmlElement {
  CT_HdrFtrRef(super.element);
  static XmlElement create(String qnTagName,
      {required WD_HEADER_FOOTER type, required String rId}) {
    final attrs = <String, String>{
      qn('w:type'): wdHeaderFooterIndexConverter.toXml(type)!,
      qn('r:id'): rId,
    };
    return OxmlElement(qnTagName, attrs: attrs);
  }

  static final qnHeaderReference = qn('w:headerReference');
  static final qnFooterReference = qn('w:footerReference');

  WD_HEADER_FOOTER get type =>
      getReqAttrVal('w:type', wdHeaderFooterIndexConverter);
  set type(WD_HEADER_FOOTER value) =>
      setReqAttrVal('w:type', value, wdHeaderFooterIndexConverter);
  WD_HEADER_FOOTER get type_ => type;
  set type_(WD_HEADER_FOOTER value) => type = value;

  String get rId => getReqAttrVal('r:id', xsdStringConverter);
  set rId(String value) => setReqAttrVal('r:id', value, xsdStringConverter);
}

/// `<w:pgMar>` element, defining page margins.
class CT_PageMar extends BaseOxmlElement {
  CT_PageMar(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:pgMar');

  Length? get top => getAttrVal('w:top', stSignedTwipsMeasureConverter);
  set top(Length? value) =>
      setAttrVal('w:top', value, stSignedTwipsMeasureConverter);

  Length? get right => getAttrVal('w:right', stTwipsMeasureConverter);
  set right(Length? value) =>
      setAttrVal('w:right', value, stTwipsMeasureConverter);

  Length? get bottom => getAttrVal('w:bottom', stSignedTwipsMeasureConverter);
  set bottom(Length? value) =>
      setAttrVal('w:bottom', value, stSignedTwipsMeasureConverter);

  Length? get left => getAttrVal('w:left', stTwipsMeasureConverter);
  set left(Length? value) =>
      setAttrVal('w:left', value, stTwipsMeasureConverter);

  Length? get header => getAttrVal('w:header', stTwipsMeasureConverter);
  set header(Length? value) =>
      setAttrVal('w:header', value, stTwipsMeasureConverter);

  Length? get footer => getAttrVal('w:footer', stTwipsMeasureConverter);
  set footer(Length? value) =>
      setAttrVal('w:footer', value, stTwipsMeasureConverter);

  Length? get gutter => getAttrVal('w:gutter', stTwipsMeasureConverter);
  set gutter(Length? value) =>
      setAttrVal('w:gutter', value, stTwipsMeasureConverter);
}

/// `<w:pgSz>` element, defining page dimensions and orientation.
class CT_PageSz extends BaseOxmlElement {
  CT_PageSz(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:pgSz');

  Length? get w => getAttrVal('w:w', stTwipsMeasureConverter);
  set w(Length? value) => setAttrVal('w:w', value, stTwipsMeasureConverter);

  Length? get h => getAttrVal('w:h', stTwipsMeasureConverter);
  set h(Length? value) => setAttrVal('w:h', value, stTwipsMeasureConverter);

  WD_ORIENTATION get orient => getAttrVal('w:orient', wdOrientationConverter,
      defaultValue: WD_ORIENTATION.PORTRAIT)!;
  set orient(WD_ORIENTATION? value) =>
      setAttrVal('w:orient', value, wdOrientationConverter,
          defaultValue: WD_ORIENTATION.PORTRAIT);
}

/// `<w:sectType>` element, defining the section start type.
class CT_SectType extends BaseOxmlElement {
  CT_SectType(super.element);
  static XmlElement create({WD_SECTION_START? val}) {
    final attrs = <String, String>{};
    if (val != null) {
      attrs[qn('w:val')] = wdSectionStartConverter.toXml(val)!;
    }
    return OxmlElement(qnTagName, attrs: attrs);
  }

  static final qnTagName = qn('w:type');

  WD_SECTION_START? get val => getAttrVal('w:val', wdSectionStartConverter);
  set val(WD_SECTION_START? value) =>
      setAttrVal('w:val', value, wdSectionStartConverter);
}

/// `<w:sectPr>` element, the container element for section properties.
class CT_SectPr extends BaseOxmlElement {
  CT_SectPr(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:sectPr');

  static final _tagSeq = [
    qn("w:footnotePr"),
    qn("w:endnotePr"),
    qn("w:type"),
    qn("w:pgSz"),
    qn("w:pgMar"),
    qn("w:paperSrc"),
    qn("w:pgBorders"),
    qn("w:lnNumType"),
    qn("w:pgNumType"),
    qn("w:cols"),
    qn("w:formProt"),
    qn("w:vAlign"),
    qn("w:noEndnote"),
    qn("w:titlePg"),
    qn("w:textDirection"),
    qn("w:bidi"),
    qn("w:rtlGutter"),
    qn("w:docGrid"),
    qn("w:printerSettings"),
    qn("w:sectPrChange"),
  ];

  static final _headerReference = ZeroOrMore<CT_HdrFtrRef>(
      qn('w:headerReference'), (el) => CT_HdrFtrRef(el));
  static final _footerReference = ZeroOrMore<CT_HdrFtrRef>(
      qn('w:footerReference'), (el) => CT_HdrFtrRef(el));
  static final _type =
      ZeroOrOne<CT_SectType>(qn('w:type'), successors: _tagSeq.sublist(3));
  static final _pgSz =
      ZeroOrOne<CT_PageSz>(qn('w:pgSz'), successors: _tagSeq.sublist(4));
  static final _pgMar =
      ZeroOrOne<CT_PageMar>(qn('w:pgMar'), successors: _tagSeq.sublist(5));
  static final _titlePg =
      ZeroOrOne<CT_OnOff>(qn('w:titlePg'), successors: _tagSeq.sublist(14));

  List<CT_HdrFtrRef> get headerReference => _headerReference.getElements(this);
  List<CT_HdrFtrRef> get footerReference => _footerReference.getElements(this);

  CT_SectType? get type => _type.getElement(this, (el) => CT_SectType(el));
  CT_PageSz? get pgSz => _pgSz.getElement(this, (el) => CT_PageSz(el));
  CT_PageMar? get pgMar => _pgMar.getElement(this, (el) => CT_PageMar(el));
  CT_OnOff? get titlePg => _titlePg.getElement(this, (el) => CT_OnOff(el));

  CT_PageMar getOrAddPgMar() =>
      _pgMar.getOrAdd(this, CT_PageMar.create, (el) => CT_PageMar(el));
  CT_PageSz getOrAddPgSz() =>
      _pgSz.getOrAdd(this, CT_PageSz.create, (el) => CT_PageSz(el));
  CT_OnOff getOrAddTitlePg() => _titlePg.getOrAdd(
      this, () => CT_OnOff.create(qn('w:titlePg')), (el) => CT_OnOff(el));
  CT_SectType getOrAddType() =>
      _type.getOrAdd(this, CT_SectType.create, (el) => CT_SectType(el));

  void removeTitlePg() => _titlePg.remove(this);
  void removeType() => _type.remove(this);
  void remove_titlePg() => removeTitlePg();
  void remove_type() => removeType();

  CT_HdrFtrRef addFooterReference(WD_HEADER_FOOTER type, String rId) {
    final qnName = CT_HdrFtrRef.qnFooterReference;
    final footerRefElement = CT_HdrFtrRef.create(qnName, type: type, rId: rId);
    insertChild(footerRefElement, _tagSeq);
    return CT_HdrFtrRef(footerRefElement);
  }

  CT_HdrFtrRef add_footerReference(WD_HEADER_FOOTER type, String rId) =>
      addFooterReference(type, rId);

  CT_HdrFtrRef addHeaderReference(WD_HEADER_FOOTER type, String rId) {
    final qnName = CT_HdrFtrRef.qnHeaderReference;
    final headerRefElement = CT_HdrFtrRef.create(qnName, type: type, rId: rId);
    insertChild(headerRefElement, _tagSeq);
    return CT_HdrFtrRef(headerRefElement);
  }

  CT_HdrFtrRef add_headerReference(WD_HEADER_FOOTER type, String rId) =>
      addHeaderReference(type, rId);

  CT_HdrFtrRef? getFooterReference(WD_HEADER_FOOTER type) {
    final typeStr = wdHeaderFooterIndexConverter.toXml(type);
    return footerReference
        .firstWhereOrNull((ref) => ref.type.xmlValue == typeStr);
  }

  CT_HdrFtrRef? get_footerReference(WD_HEADER_FOOTER type) =>
      getFooterReference(type);

  CT_HdrFtrRef? getHeaderReference(WD_HEADER_FOOTER type) {
    final typeStr = wdHeaderFooterIndexConverter.toXml(type);
    return headerReference
        .firstWhereOrNull((ref) => ref.type.xmlValue == typeStr);
  }

  CT_HdrFtrRef? get_headerReference(WD_HEADER_FOOTER type) =>
      getHeaderReference(type);

  String removeFooterReference(WD_HEADER_FOOTER type) {
    final footerRef = getFooterReference(type);
    if (footerRef == null) {
      throw ArgumentError("CT_SectPr has no footer reference of type $type");
    }
    final rId = footerRef.rId;
    element.children.remove(footerRef.element);
    return rId;
  }

  String remove_footerReference(WD_HEADER_FOOTER type) =>
      removeFooterReference(type);

  String removeHeaderReference(WD_HEADER_FOOTER type) {
    final headerRef = getHeaderReference(type);
    if (headerRef == null) {
      throw ArgumentError("CT_SectPr has no header reference of type $type");
    }
    final rId = headerRef.rId;
    element.children.remove(headerRef.element);
    return rId;
  }

  String remove_headerReference(WD_HEADER_FOOTER type) =>
      removeHeaderReference(type);

  Length? get bottomMargin => pgMar?.bottom;
  set bottomMargin(Length? value) => getOrAddPgMar().bottom = value;
  Length? get bottom_margin => bottomMargin;
  set bottom_margin(Length? value) => bottomMargin = value;

  Length? get footerDistance => pgMar?.footer;
  set footerDistance(Length? value) => getOrAddPgMar().footer = value;
  Length? get footer => footerDistance;
  set footer(Length? value) => footerDistance = value;

  Length? get gutter => pgMar?.gutter;
  set gutter(Length? value) => getOrAddPgMar().gutter = value;

  Length? get headerDistance => pgMar?.header;
  set headerDistance(Length? value) => getOrAddPgMar().header = value;
  Length? get header => headerDistance;
  set header(Length? value) => headerDistance = value;

  Length? get leftMargin => pgMar?.left;
  set leftMargin(Length? value) => getOrAddPgMar().left = value;
  Length? get left_margin => leftMargin;
  set left_margin(Length? value) => leftMargin = value;

  Length? get rightMargin => pgMar?.right;
  set rightMargin(Length? value) => getOrAddPgMar().right = value;
  Length? get right_margin => rightMargin;
  set right_margin(Length? value) => rightMargin = value;

  Length? get topMargin => pgMar?.top;
  set topMargin(Length? value) => getOrAddPgMar().top = value;
  Length? get top_margin => topMargin;
  set top_margin(Length? value) => topMargin = value;

  WD_ORIENTATION get orientation => pgSz?.orient ?? WD_ORIENTATION.PORTRAIT;
  set orientation(WD_ORIENTATION? value) =>
      getOrAddPgSz().orient = value ?? WD_ORIENTATION.PORTRAIT;

  Length? get pageHeight => pgSz?.h;
  set pageHeight(Length? value) => getOrAddPgSz().h = value;
  Length? get page_height => pageHeight;
  set page_height(Length? value) => pageHeight = value;

  Length? get pageWidth => pgSz?.w;
  set pageWidth(Length? value) => getOrAddPgSz().w = value;
  Length? get page_width => pageWidth;
  set page_width(Length? value) => pageWidth = value;

  WD_SECTION_START get startType => type?.val ?? WD_SECTION_START.NEW_PAGE;
  set startType(WD_SECTION_START? value) {
    if (value == null || value == WD_SECTION_START.NEW_PAGE) {
      removeType();
    } else {
      getOrAddType().val = value;
    }
  }

  WD_SECTION_START get start_type => startType;
  set start_type(WD_SECTION_START? value) => startType = value;

  bool get titlePgVal => titlePg?.val ?? false;
  set titlePgVal(bool? value) {
    if (value == null || !value) {
      removeTitlePg();
    } else {
      getOrAddTitlePg().val = true;
    }
  }

  bool get titlePg_val => titlePgVal;
  set titlePg_val(bool? value) => titlePgVal = value;

  CT_SectPr clone() {
    final clonedElement = element.copy(); // Assure it's XmlElement
    clonedElement.attributes.removeWhere((attr) =>
        attr.name.prefix == 'w' && attr.name.local.startsWith('rsid'));
    return CT_SectPr(clonedElement);
  }

  CT_SectPr? get precedingSectPr {
    print('Warn: CT_SectPr.precedingSectPr is not fully implemented.');
    return null;
  }

  CT_SectPr? get preceding_sectPr => precedingSectPr;

  Iterable<BaseOxmlElement /* CT_P | CT_Tbl */ > iterInnerContent() sync* {
    print(
        'Warn: CT_SectPr.iterInnerContent is not fully implemented and will yield no results.');
    yield* [];
  }

  Iterable<BaseOxmlElement /* CT_P | CT_Tbl */ > iter_inner_content() =>
      iterInnerContent();
}

// Placeholder classes for required converters (Assume defined in simpletypes.dart)
class WD_HEADER_FOOTER_Converter implements BaseSimpleType<WD_HEADER_FOOTER> {
  const WD_HEADER_FOOTER_Converter();
  @override
  WD_HEADER_FOOTER fromXml(String v) => WD_HEADER_FOOTER.fromXml(v);
  @override
  String? toXml(v) => v.xmlValue;
  @override
  void validate(v) {}
}

class WD_ORIENTATION_Converter implements BaseSimpleType<WD_ORIENTATION> {
  const WD_ORIENTATION_Converter();
  @override
  WD_ORIENTATION fromXml(String v) => WD_ORIENTATION.fromXml(v);
  @override
  String? toXml(v) => v.xmlValue;
  @override
  void validate(v) {}
}

class WD_SECTION_START_Converter implements BaseSimpleType<WD_SECTION_START> {
  const WD_SECTION_START_Converter();
  @override
  WD_SECTION_START fromXml(String v) => WD_SECTION_START.fromXml(v);
  @override
  String? toXml(v) => v.xmlValue;
  @override
  void validate(v) {}
}

class XsdStringConverter implements BaseSimpleType<String> {
  const XsdStringConverter();
  @override
  String fromXml(String v) => v;
  @override
  String? toXml(v) => v;
  @override
  void validate(v) {}
}

class ST_SignedTwipsMeasureConverter implements BaseSimpleType<Length> {
  const ST_SignedTwipsMeasureConverter();
  @override
  Length fromXml(String v) => Twips(int.parse(v));
  @override
  String? toXml(v) => v.twips.toString();
  @override
  void validate(v) {}
}

class ST_TwipsMeasureConverter implements BaseSimpleType<Length> {
  const ST_TwipsMeasureConverter();
  @override
  Length fromXml(String v) => Twips(int.parse(v));
  @override
  String? toXml(v) => v.twips.toString();
  @override
  void validate(v) {
    if (v.emu < 0) throw ArgumentError('must be non-negative');
  }
}

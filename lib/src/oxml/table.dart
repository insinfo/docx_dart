/// Path: lib/src/oxml/table.dart
/// Based on python-docx: docx/oxml/table.py
/// Custom element classes for tables (<w:tbl>, <w:tr>, <w:tc>, etc.).

import 'package:docx_dart/src/oxml/shared.dart';
import 'package:xml/xml.dart';

import '../enum/table.dart' // Enums específicos de tabela
  show
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
    WD_TABLE_DIRECTION;

import '../enum/text.dart';

import '../shared.dart' show Emu, Length, Twips;
import 'ns.dart';
import 'parser.dart';

import 'simpletypes.dart';
// Importa CT_P de text/paragraph
import 'text/paragraph.dart' show CT_P;

import 'text/parfmt.dart' show CT_Jc;
import 'xmlchemy.dart' show BaseOxmlElement;

import 'xmlchemy_descriptors.dart';

// ignore_for_file: camel_case_types, unnecessary_this, unused_field

// --- Conversores Globais (Definidos e Instanciados AQUI ou importados de simpletypes.dart) ---
// É MELHOR definir estas classes em simpletypes.dart e apenas instanciar/importar aqui.
// Assumindo que estão definidos em simpletypes.dart e BaseSimpleType também.

// Placeholder classes (REMOVA se definidas em simpletypes.dart)
class ST_TwipsMeasureConverter implements BaseSimpleType<Length> {
  const ST_TwipsMeasureConverter();
  @override
  Length fromXml(String xmlValue) => Twips(int.parse(xmlValue));
  @override
  String? toXml(Length? value) => value?.twips.toString();
  @override
  void validate(Length value) {
    if (value.emu < 0) throw ArgumentError("TwipsMeasure must be non-negative");
  }
}

class WD_ROW_HEIGHT_RULE_Converter
    implements BaseSimpleType<WD_ROW_HEIGHT_RULE> {
  const WD_ROW_HEIGHT_RULE_Converter();
  @override
  WD_ROW_HEIGHT_RULE fromXml(String xmlValue) => WD_ROW_HEIGHT_RULE
      .fromXml(xmlValue); // Assume static fromXml exists on Enum
  @override
  String? toXml(WD_ROW_HEIGHT_RULE? value) =>
      value?.xmlValue; // Assume xmlValue exists on Enum
  @override
  void validate(WD_ROW_HEIGHT_RULE value) {}
}

class ST_TblLayoutTypeConverter implements BaseSimpleType<String> {
  const ST_TblLayoutTypeConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {
    if (value != 'fixed' && value != 'autofit')
      throw ArgumentError("ST_TblLayoutType must be 'fixed' or 'autofit'");
  }
}

class ST_TblWidthConverter implements BaseSimpleType<String> {
  const ST_TblWidthConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {
    if (!['dxa', 'pct', 'auto', 'nil'].contains(value))
      throw ArgumentError("ST_TblWidth invalid");
  }
}

class XsdIntConverter implements BaseSimpleType<int> {
  const XsdIntConverter();
  @override
  int fromXml(String xmlValue) => int.parse(xmlValue);
  @override
  String? toXml(int? value) => value?.toString();
  @override
  void validate(int value) {}
}

class WD_CELL_VERTICAL_ALIGNMENT_Converter
    implements BaseSimpleType<WD_CELL_VERTICAL_ALIGNMENT> {
  const WD_CELL_VERTICAL_ALIGNMENT_Converter();
  @override
  WD_CELL_VERTICAL_ALIGNMENT fromXml(String xmlValue) =>
      WD_CELL_VERTICAL_ALIGNMENT.fromXml(xmlValue);
  @override
  String? toXml(WD_CELL_VERTICAL_ALIGNMENT? value) => value?.xmlValue;
  @override
  void validate(WD_CELL_VERTICAL_ALIGNMENT value) {}
}

class ST_MergeConverter implements BaseSimpleType<String> {
  const ST_MergeConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {
    if (value != ST_Merge.CONTINUE && value != ST_Merge.RESTART)
      throw ArgumentError("ST_Merge must be 'continue' or 'restart'");
  }
}

// Converter para o alinhamento da tabela (usado em CT_TblPr.alignment)
class WD_TABLE_ALIGNMENT_Converter
    implements BaseSimpleType<WD_TABLE_ALIGNMENT> {
  const WD_TABLE_ALIGNMENT_Converter();
  @override
  WD_TABLE_ALIGNMENT fromXml(String xmlValue) =>
      WD_TABLE_ALIGNMENT.fromXml(xmlValue);
  @override
  String? toXml(WD_TABLE_ALIGNMENT? value) => value?.xmlValue;
  @override
  void validate(WD_TABLE_ALIGNMENT value) {}
}

// Assume ST_Merge constants exist in simpletypes.dart or locally
class ST_Merge {
  static const CONTINUE = 'continue';
  static const RESTART = 'restart';
}

// Instanciando os conversores
final stTwipsMeasureConverter = const ST_TwipsMeasureConverter();
final wdRowHeightRuleConverter = const WD_ROW_HEIGHT_RULE_Converter();
final stTblLayoutTypeConverter = const ST_TblLayoutTypeConverter();
final stTblWidthConverter = const ST_TblWidthConverter();
final xsdIntConverter = const XsdIntConverter();
final wdCellVerticalAlignmentConverter =
    const WD_CELL_VERTICAL_ALIGNMENT_Converter();
final stMergeConverter = const ST_MergeConverter();
final wdTableAlignmentConverter =
    const WD_TABLE_ALIGNMENT_Converter(); // Instancia para alinhamento da tabela
// --- Fim dos Conversores ---

/// `<w:trHeight>` element, specifying row height properties.
class CT_Height extends BaseOxmlElement {
  CT_Height(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:trHeight');

  Length? get val => getAttrVal('w:val', stTwipsMeasureConverter);
  set val(Length? value) => setAttrVal('w:val', value, stTwipsMeasureConverter);

  WD_ROW_HEIGHT_RULE? get hRule =>
      getAttrVal('w:hRule', wdRowHeightRuleConverter);
  set hRule(WD_ROW_HEIGHT_RULE? value) =>
      setAttrVal('w:hRule', value, wdRowHeightRuleConverter);
}

/// `<w:tblPrEx>` element, table property exceptions.
class CT_TblPrEx extends BaseOxmlElement {
  CT_TblPrEx(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:tblPrEx');
}

/// `<w:trPr>` element, defining table row properties.
class CT_TrPr extends BaseOxmlElement {
  CT_TrPr(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:trPr');

  static final _childSequence = [
    qn("w:cnfStyle"),
    qn("w:divId"),
    qn("w:gridBefore"),
    qn("w:gridAfter"),
    qn("w:wBefore"),
    qn("w:wAfter"),
    qn("w:cantSplit"),
    qn("w:trHeight"),
    qn("w:tblHeader"),
    qn("w:tblCellSpacing"),
    qn("w:jc"),
    qn("w:hidden"),
    qn("w:ins"),
    qn("w:del"),
    qn("w:trPrChange"),
  ];

  static final _gridAfter = ZeroOrOne<CT_DecimalNumber>(qn('w:gridAfter'),
      successors: _childSequence.sublist(4));
  static final _gridBefore = ZeroOrOne<CT_DecimalNumber>(qn('w:gridBefore'),
      successors: _childSequence.sublist(3));
  static final _trHeight = ZeroOrOne<CT_Height>(qn('w:trHeight'),
      successors: _childSequence.sublist(8));

  int get gridAfter =>
      _gridAfter.getElement(this, (el) => CT_DecimalNumber(el))?.val ?? 0;
  int get gridBefore =>
      _gridBefore.getElement(this, (el) => CT_DecimalNumber(el))?.val ?? 0;

  CT_Height? get trHeightElem =>
      _trHeight.getElement(this, (el) => CT_Height(el));
  CT_Height getOrAddTrHeight() =>
      _trHeight.getOrAdd(this, CT_Height.create, (el) => CT_Height(el));

  WD_ROW_HEIGHT_RULE? get trHeight_hRule => trHeightElem?.hRule;
  set trHeight_hRule(WD_ROW_HEIGHT_RULE? value) {
    if (value == null && trHeightElem == null) return;
    getOrAddTrHeight().hRule = value;
  }

  Length? get trHeight_val => trHeightElem?.val;
  set trHeight_val(Length? value) {
    if (value == null && trHeightElem == null) return;
    getOrAddTrHeight().val = value;
  }
}

/// `<w:tcPr>` element, defining table cell properties.
class CT_TcPr extends BaseOxmlElement {
  CT_TcPr(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:tcPr');

  static final _childSequence = [
    qn("w:cnfStyle"),
    qn("w:tcW"),
    qn("w:gridSpan"),
    qn("w:hMerge"),
    qn("w:vMerge"),
    qn("w:tcBorders"),
    qn("w:shd"),
    qn("w:noWrap"),
    qn("w:tcMar"),
    qn("w:textDirection"),
    qn("w:tcFitText"),
    qn("w:vAlign"),
    qn("w:hideMark"),
    qn("w:headers"),
    qn("w:cellIns"),
    qn("w:cellDel"),
    qn("w:cellMerge"),
    qn("w:tcPrChange"),
  ];

  static final _tcW = ZeroOrOne<CT_TblWidth>(qn('w:tcW'),
      successors: _childSequence.sublist(2));
  static final _gridSpan = ZeroOrOne<CT_DecimalNumber>(qn('w:gridSpan'),
      successors: _childSequence.sublist(3));
  static final _vMerge = ZeroOrOne<CT_VMerge>(qn('w:vMerge'),
      successors: _childSequence.sublist(5));
  static final _vAlign = ZeroOrOne<CT_VerticalJc>(qn('w:vAlign'),
      successors: _childSequence.sublist(12));

  CT_TblWidth? get tcW => _tcW.getElement(this, (el) => CT_TblWidth(el));
  CT_TblWidth getOrAddTcW() =>
      _tcW.getOrAdd(this, CT_TblWidth.create, (el) => CT_TblWidth(el));

  CT_DecimalNumber? get gridSpanElement =>
      _gridSpan.getElement(this, (el) => CT_DecimalNumber(el));
  CT_DecimalNumber getOrAddGridSpan() => _gridSpan.getOrAdd(
      this,
      () => CT_DecimalNumber.create(qn('w:gridSpan'), 0),
      (el) => CT_DecimalNumber(el));

  CT_VMerge? get vMergeElement =>
      _vMerge.getElement(this, (el) => CT_VMerge(el));
  CT_VMerge getOrAddVMerge() =>
      _vMerge.getOrAdd(this, CT_VMerge.create, (el) => CT_VMerge(el));

  CT_VerticalJc? get vAlignElement =>
      _vAlign.getElement(this, (el) => CT_VerticalJc(el));
  CT_VerticalJc getOrAddVAlign() => _vAlign.getOrAdd(
      this,
      () => CT_VerticalJc.create(val: WD_CELL_VERTICAL_ALIGNMENT.TOP),
      (el) => CT_VerticalJc(el));

  int get gridSpan => gridSpanElement?.val ?? 1;
  set gridSpan(int value) {
    if (value <= 1) {
      _gridSpan.remove(this);
    } else {
      getOrAddGridSpan().val = value;
    }
  }

  String? get vMergeVal => vMergeElement?.val;
  set vMergeVal(String? newVal) {
    if (newVal == null) {
      _vMerge.remove(this);
    } else {
      getOrAddVMerge().val = newVal;
    }
  }

  WD_CELL_VERTICAL_ALIGNMENT? get vAlignVal => vAlignElement?.val;
  set vAlignVal(WD_CELL_VERTICAL_ALIGNMENT? v) {
    if (v == null) {
      _vAlign.remove(this);
    } else {
      getOrAddVAlign().val = v;
    }
  }

  Length? get width => tcW?.width;
  set width(Length? val) {
    if (val == null) {
      _tcW.remove(this);
    } else {
      getOrAddTcW().width = val;
    }
  }
}

/// `<w:tc>` table cell element.
class CT_Tc extends BaseOxmlElement {
  CT_Tc(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent, then add child ---
    final tc = OxmlElement(qnTagName);
    tc.children.add(CT_P.create()); // Add required <w:p>
    return tc;
    // --- End Correction ---
  }

  static final qnTagName = qn('w:tc');

  static final _tcPr =
      ZeroOrOne<CT_TcPr>(qn('w:tcPr'), successors: [qn('w:p'), qn('w:tbl')]);
  static final _p = OneOrMore<CT_P>(qn('w:p'), (el) => CT_P(el));
  static final _tbl = ZeroOrMore<CT_Tbl>(qn('w:tbl'), (el) => CT_Tbl(el));

  CT_TcPr? get tcPr => _tcPr.getElement(this, (el) => CT_TcPr(el));
  CT_TcPr getOrAddTcPr() =>
      _tcPr.getOrAdd(this, CT_TcPr.create, (el) => CT_TcPr(el));

  List<CT_P> get pList => _p.getElements(this);
  List<CT_Tbl> get tblList => _tbl.getElements(this);
  List<BaseOxmlElement> get innerContentElements {
    final content = <BaseOxmlElement>[];
    for (final child in element.children.whereType<XmlElement>()) {
      final nsUri = child.name.namespaceUri;
      final localName = child.name.local;
      if (nsUri != nsmap['w']) {
        continue;
      }
      if (localName == 'p') {
        content.add(CT_P(child));
      } else if (localName == 'tbl') {
        content.add(CT_Tbl(child));
      }
    }
    return content;
  }

  int get gridSpan => tcPr?.gridSpan ?? 1;
  set gridSpan(int value) => getOrAddTcPr().gridSpan = value;
  int get grid_span => gridSpan;
  set grid_span(int value) => gridSpan = value;

  String? get vMerge => tcPr?.vMergeVal;
  set vMerge(String? val) => getOrAddTcPr().vMergeVal = val;

  Length? get width => tcPr?.width;
  set width(Length? val) => getOrAddTcPr().width = val;

  int get gridOffset {
    final row = _tr;
    var offset = row.gridBefore;
    for (final cell in row.tcList) {
      if (cell.element == element) {
        return offset;
      }
      offset += cell.gridSpan;
    }
    return offset;
  }

  int get grid_offset => gridOffset;

  List<BaseOxmlElement> get inner_content_elements => innerContentElements;

  void clearContent() {
    final removable = <XmlNode>[];
    for (final child in element.children) {
      if (child is XmlElement) {
        final isTcPr =
            child.name.local == 'tcPr' && child.name.namespaceUri == nsmap['w'];
        if (isTcPr) {
          continue;
        }
      }
      removable.add(child);
    }
    for (final node in removable) {
      element.children.remove(node);
    }
  }

  void clear_content() => clearContent();

  CT_P addP() {
    final pElement = CT_P.create();
    element.children.add(pElement);
    return CT_P(pElement);
  }

  CT_Row get _tr {
    final parentRow = getParentAs<CT_Row>(CT_Row.new);
    if (parentRow == null) {
      throw StateError('CT_Tc has no parent row');
    }
    return parentRow;
  }

  CT_Tbl get _enclosingTbl {
    final tbl = _tr.getParentAs<CT_Tbl>(CT_Tbl.new);
    if (tbl == null) {
      throw StateError('CT_Tc has no ancestor table');
    }
    return tbl;
  }

  int get _trIdx {
    final rows = _enclosingTbl.trList;
    final idx = rows.indexWhere((row) => row.element == _tr.element);
    if (idx < 0) {
      throw StateError('Row for CT_Tc not found in parent table');
    }
    return idx;
  }

  CT_Row get _trAbove {
    final rows = _enclosingTbl.trList;
    final idx = _trIdx;
    if (idx == 0) {
      throw StateError('Top-most row has no row above');
    }
    return rows[idx - 1];
  }

  CT_Row? get _trBelow {
    final rows = _enclosingTbl.trList;
    final idx = _trIdx;
    if (idx >= rows.length - 1) {
      return null;
    }
    return rows[idx + 1];
  }

  CT_Tc get _tcAbove => _trAbove.tcAtGridOffset(gridOffset);
  // ignore: unused_element
  CT_Tc? get _tcBelow {
    final rowBelow = _trBelow;
    if (rowBelow == null) {
      return null;
    }
    return rowBelow.tcAtGridOffset(gridOffset);
  }

  CT_Tc get tcAbove => _tcAbove;

  CT_Tc merge(CT_Tc otherTc) {
    print('WARN: CT_Tc.merge() is not fully implemented.');
    throw UnimplementedError(
        'CT_Tc.merge() requires complex table structure manipulation.');
  }
}

/// `<w:tr>` element.
class CT_Row extends BaseOxmlElement {
  CT_Row(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:tr');

  static final _childSequence = [
    CT_TblPrEx.qnTagName,
    CT_TrPr.qnTagName,
    CT_Tc.qnTagName
  ];

  static final _tblPrEx = ZeroOrOne<CT_TblPrEx>(qn('w:tblPrEx'),
      successors: _childSequence.sublist(1));
  static final _trPr =
      ZeroOrOne<CT_TrPr>(qn('w:trPr'), successors: _childSequence.sublist(2));
  static final _tc = ZeroOrMore<CT_Tc>(qn('w:tc'), (el) => CT_Tc(el));

  CT_TblPrEx? get tblPrEx => _tblPrEx.getElement(this, (el) => CT_TblPrEx(el));
  CT_TrPr? get trPr => _trPr.getElement(this, (el) => CT_TrPr(el));
  CT_TrPr getOrAddTrPr() =>
      _trPr.getOrAdd(this, CT_TrPr.create, (el) => CT_TrPr(el));

  List<CT_Tc> get tcList => _tc.getElements(this);
  List<CT_Tc> get tc_lst => tcList;

  int get gridAfter => trPr?.gridAfter ?? 0;
  int get gridBefore => trPr?.gridBefore ?? 0;

  int get trIdx {
    final tbl = getParentAs<CT_Tbl>(CT_Tbl.new);
    if (tbl == null) return -1;
    return tbl.trList.indexWhere((row) => row.element == this.element);
  }

  int get tr_idx => trIdx;

  WD_ROW_HEIGHT_RULE? get trHeight_hRule => trPr?.trHeight_hRule;
  set trHeight_hRule(WD_ROW_HEIGHT_RULE? val) =>
      getOrAddTrPr().trHeight_hRule = val;

  Length? get trHeight_val => trPr?.trHeight_val;
  set trHeight_val(Length? val) => getOrAddTrPr().trHeight_val = val;

  CT_Tc addTc() {
    final tcElement = CT_Tc.create();
    element.children.add(tcElement);
    return CT_Tc(tcElement);
  }

  CT_Tc add_tc() => addTc();

  CT_Tc tcAtGridOffset(int gridOffset) {
    int remainingOffset = gridOffset - gridBefore;
    if (remainingOffset < 0) {
      throw ArgumentError(
          "grid_offset=$gridOffset falls within gridBefore region ($gridBefore)");
    }
    for (final cell in tcList) {
      if (remainingOffset == 0) return cell;
      if (remainingOffset < 0) break;
      remainingOffset -= cell.gridSpan;
    }
    throw ArgumentError(
        "no `tc` element starts exactly at grid_offset=$gridOffset in this row");
  }

  CT_Tc tc_at_grid_offset(int grid_offset) => tcAtGridOffset(grid_offset);
}

/// `<w:gridCol>` element, defines a table column in the grid.
class CT_TblGridCol extends BaseOxmlElement {
  CT_TblGridCol(super.element);
  static XmlElement create({Length? w}) {
    final attrs = <String, String>{};
    if (w != null) attrs[qn('w:w')] = stTwipsMeasureConverter.toXml(w)!;
    return OxmlElement(qnTagName, attrs: attrs);
  }

  static final qnTagName = qn('w:gridCol');

  Length? get w => getAttrVal('w:w', stTwipsMeasureConverter);
  set w(Length? value) => setAttrVal('w:w', value, stTwipsMeasureConverter);

  int get gridColIdx {
    final grid = getParentAs<CT_TblGrid>(CT_TblGrid.new);
    if (grid == null) return -1;
    return grid.gridColList.indexWhere((col) => col.element == this.element);
  }

  int get gridCol_idx => gridColIdx;
}

/// `<w:tblGrid>` element, container for `<w:gridCol>` elements.
class CT_TblGrid extends BaseOxmlElement {
  CT_TblGrid(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:tblGrid');

  static final _gridCol =
      ZeroOrMore<CT_TblGridCol>(qn('w:gridCol'), (el) => CT_TblGridCol(el));

  List<CT_TblGridCol> get gridColList => _gridCol.getElements(this);
  List<CT_TblGridCol> get gridCol_lst => gridColList;

  CT_TblGridCol addGridCol() {
    final gridColElement = CT_TblGridCol.create();
    element.children.add(gridColElement);
    return CT_TblGridCol(gridColElement);
  }

  CT_TblGridCol add_gridCol() => addGridCol();
}

/// `<w:tblLayout>` element, specifies table layout type (fixed/autofit).
class CT_TblLayoutType extends BaseOxmlElement {
  CT_TblLayoutType(super.element);
  static XmlElement create({String type = 'autofit'}) =>
      OxmlElement(qnTagName, attrs: {qn('w:type'): type});
  static final qnTagName = qn('w:tblLayout');

  String? get type => getAttrVal('w:type', stTblLayoutTypeConverter);
  set type(String? value) =>
      setAttrVal('w:type', value, stTblLayoutTypeConverter);
}

/// `<w:tblW>` element, specifies table-related width.
class CT_TblW extends BaseOxmlElement {
  CT_TblW(super.element);
  static XmlElement create({int w = 0, String type = 'dxa'}) =>
      OxmlElement(qnTagName,
          attrs: {qn('w:w'): w.toString(), qn('w:type'): type});
  static final qnTagName = qn('w:tblW');
  static final qnTcw = qn('w:tcW');

  int get w => getReqAttrVal('w:w', xsdIntConverter);
  set w(int value) => setReqAttrVal('w:w', value, xsdIntConverter);

  String get type => getReqAttrVal('w:type', stTblWidthConverter);
  set type(String value) => setReqAttrVal('w:type', value, stTblWidthConverter);

  Length? get width {
    if (type != 'dxa') return null;
    return Twips(w);
  }

  set width(Length? value) {
    if (value == null) {
      type = 'auto';
      w = 0;
    } else {
      type = 'dxa';
      w = value.twips;
    }
  }
}

/// `<w:tblPr>` element, child of `<w:tbl>`, holds table properties.
class CT_TblPr extends BaseOxmlElement {
  CT_TblPr(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:tblPr');

  static final _childSequence = [
    qn("w:tblStyle"),
    qn("w:tblpPr"),
    qn("w:tblOverlap"),
    qn("w:bidiVisual"),
    qn("w:tblStyleRowBandSize"),
    qn("w:tblStyleColBandSize"),
    qn("w:tblW"),
    qn("w:jc"),
    qn("w:tblCellSpacing"),
    qn("w:tblInd"),
    qn("w:tblBorders"),
    qn("w:shd"),
    qn("w:tblLayout"),
    qn("w:tblCellMar"),
    qn("w:tblLook"),
    qn("w:tblCaption"),
    qn("w:tblDescription"),
    qn("w:tblPrChange"),
  ];

  static final _tblStyle = ZeroOrOne<CT_String>(qn('w:tblStyle'),
      successors: _childSequence.sublist(1));
  static final _bidiVisual = ZeroOrOne<CT_OnOff>(qn('w:bidiVisual'),
      successors: _childSequence.sublist(4));
  static final _tblW =
      ZeroOrOne<CT_TblW>(qn('w:tblW'), successors: _childSequence.sublist(7));
  static final _jc =
      ZeroOrOne<CT_Jc>(qn('w:jc'), successors: _childSequence.sublist(8));
  static final _tblLayout = ZeroOrOne<CT_TblLayoutType>(qn('w:tblLayout'),
      successors: _childSequence.sublist(13));

  // --- Property Accessors ---
  CT_String? get tblStyle => _tblStyle.getElement(this, (el) => CT_String(el));
  CT_OnOff? get bidiVisual =>
      _bidiVisual.getElement(this, (el) => CT_OnOff(el));
  CT_Jc? get jc => _jc.getElement(this, (el) => CT_Jc(el));
  CT_TblLayoutType? get tblLayout =>
      _tblLayout.getElement(this, (el) => CT_TblLayoutType(el));
  CT_TblW? get tblW => _tblW.getElement(this, (el) => CT_TblW(el));

  CT_String getOrAddTblStyle() => _tblStyle.getOrAdd(this,
      () => CT_String.create(qn('w:tblStyle'), ''), (el) => CT_String(el));
  CT_OnOff getOrAddBidiVisual() => _bidiVisual.getOrAdd(
      this, () => CT_OnOff.create(qn('w:bidiVisual')), (el) => CT_OnOff(el));
  CT_Jc getOrAddJc() => _jc.getOrAdd(
      this,
      () => CT_Jc.create(WD_ALIGN_PARAGRAPH.LEFT),
      (el) => CT_Jc(el)); // Use PARAGRAPH enum here? Or TABLE?
  CT_TblLayoutType getOrAddTblLayout() => _tblLayout.getOrAdd(
      this, CT_TblLayoutType.create, (el) => CT_TblLayoutType(el));
  CT_TblW getOrAddTblW() =>
      _tblW.getOrAdd(this, CT_TblW.create, (el) => CT_TblW(el));

  // --- Higher-level properties ---

  // --- CORRECTION: Convert between WD_TABLE_ALIGNMENT and WD_PARAGRAPH_ALIGNMENT ---
  WD_TABLE_ALIGNMENT? get alignment {
    final jcElement = jc;
    if (jcElement == null) return null;
    // Convert paragraph alignment enum to table alignment enum
    switch (jcElement.val) {
      case WD_PARAGRAPH_ALIGNMENT.LEFT:
        return WD_TABLE_ALIGNMENT.LEFT;
      case WD_PARAGRAPH_ALIGNMENT.CENTER:
        return WD_TABLE_ALIGNMENT.CENTER;
      case WD_PARAGRAPH_ALIGNMENT.RIGHT:
        return WD_TABLE_ALIGNMENT.RIGHT;
      // Other WD_PARAGRAPH_ALIGNMENT values don't map directly to WD_TABLE_ALIGNMENT
      default:
        return null; // Or throw? Or map JUSTIFY? Check OOXML spec for <w:jc> within <w:tblPr>
    }
  }

  set alignment(WD_TABLE_ALIGNMENT? value) {
    if (value == null) {
      _jc.remove(this);
    } else {
      // Convert table alignment enum to paragraph alignment enum for storage in <w:jc>
      WD_PARAGRAPH_ALIGNMENT paraAlign;
      switch (value) {
        case WD_TABLE_ALIGNMENT.LEFT:
          paraAlign = WD_PARAGRAPH_ALIGNMENT.LEFT;
          break;
        case WD_TABLE_ALIGNMENT.CENTER:
          paraAlign = WD_PARAGRAPH_ALIGNMENT.CENTER;
          break;
        case WD_TABLE_ALIGNMENT.RIGHT:
          paraAlign = WD_PARAGRAPH_ALIGNMENT.RIGHT;
          break;
      }
      getOrAddJc().val = paraAlign;
    }
  }
  // --- End Correction ---

  bool get autofit => tblLayout?.type != 'fixed';
  set autofit(bool v) {
    getOrAddTblLayout().type = v ? 'autofit' : 'fixed';
  }

  String? get style => tblStyle?.val;
  set style(String? styleId) {
    if (styleId == null) {
      _tblStyle.remove(this);
    } else {
      getOrAddTblStyle().val = styleId;
    }
  }

  Length? get width => tblW?.width;
  set width(Length? value) {
    if (value == null) {
      final w = getOrAddTblW();
      w.type = 'auto';
      w.w = 0;
    } else {
      getOrAddTblW().width = value;
    }
  }

  bool? get bidiVisualVal => bidiVisual?.val;
  set bidiVisualVal(bool? v) {
    if (v == null) {
      _bidiVisual.remove(this);
    } else {
      getOrAddBidiVisual().val = v;
    }
  }
}

/// `<w:tbl>` element.
class CT_Tbl extends BaseOxmlElement {
  CT_Tbl(super.element);

  /// Creates a new minimal `<w:tbl>` element with required children `<w:tblPr>` and `<w:tblGrid>`.
  static XmlElement create() {
    // --- CORRECTION: Create parent, then add children ---
    final tbl = OxmlElement(qnTagName); // Cria <w:tbl> vazio
    tbl.children.add(CT_TblPr.create()); // Cria e adiciona <w:tblPr>
    tbl.children.add(CT_TblGrid.create()); // Cria e adiciona <w:tblGrid>
    return tbl;
    // --- End Correction ---
  }

  static final qnTagName = qn('w:tbl');

  static final _tblPr =
      OneAndOnlyOne<CT_TblPr>(qn('w:tblPr'), (el) => CT_TblPr(el));
  static final _tblGrid =
      OneAndOnlyOne<CT_TblGrid>(qn('w:tblGrid'), (el) => CT_TblGrid(el));
  static final _tr = ZeroOrMore<CT_Row>(qn('w:tr'), (el) => CT_Row(el));

  CT_TblPr get tblPr => _tblPr.getElement(this);
  CT_TblGrid get tblGrid => _tblGrid.getElement(this);
  List<CT_Row> get trList => _tr.getElements(this);
  List<CT_Row> get tr_lst => trList;

  int get colCount => tblGrid.gridColList.length;
  int get col_count => colCount;

  CT_Row addTr() {
    final trElement = CT_Row.create();
    element.children.add(trElement);
    return CT_Row(trElement);
  }

  CT_Row add_tr() => addTr();

  Iterable<CT_Tc> iterTcs() sync* {
    for (final row in trList) {
      yield* row.tcList;
    }
  }

  Iterable<CT_Tc> iter_tcs() => iterTcs();

  static CT_Tbl newTbl(int rows, int cols, Length width) {
    final tblElement = CT_Tbl.create();
    final tbl = CT_Tbl(tblElement);
    final tblPr = tbl.tblPr;
    final tblW = tblPr.getOrAddTblW();
    tblW.type = 'auto';
    tblW.w = 0;

    final tblGrid = tbl.tblGrid;
    final colWidth = (cols > 0) ? Emu(width.emu ~/ cols) : Emu(0);
    for (int i = 0; i < cols; i++) {
      final gridCol = tblGrid.addGridCol();
      gridCol.w = colWidth;
    }

    for (int r = 0; r < rows; r++) {
      final tr = tbl.addTr();
      for (int c = 0; c < cols; c++) {
        final tc = tr.addTc();
        tc.width = colWidth;
      }
    }
    return tbl;
  }

  static CT_Tbl new_tbl(int rows, int cols, Length width) =>
      newTbl(rows, cols, width);

  String? get tblStyleVal => tblPr.style;
  set tblStyleVal(String? styleId) => tblPr.style = styleId;
  String? get tblStyle_val => tblStyleVal;
  set tblStyle_val(String? styleId) => tblStyleVal = styleId;

  WD_TABLE_DIRECTION? get bidiVisualVal {
    final bool? raw = tblPr.bidiVisualVal;
    if (raw == null) {
      return null;
    }
    return raw ? WD_TABLE_DIRECTION.RTL : WD_TABLE_DIRECTION.LTR;
  }

  set bidiVisualVal(WD_TABLE_DIRECTION? direction) {
    if (direction == null) {
      tblPr.bidiVisualVal = null;
    } else {
      tblPr.bidiVisualVal = direction == WD_TABLE_DIRECTION.RTL;
    }
  }

  WD_TABLE_DIRECTION? get bidiVisual_val => bidiVisualVal;
  set bidiVisual_val(WD_TABLE_DIRECTION? direction) =>
      bidiVisualVal = direction;
}

/// `<w:tblW>` element, specifies table-related width.
class CT_TblWidth extends BaseOxmlElement {
  CT_TblWidth(super.element);
  static XmlElement create({int w = 0, String type = 'dxa'}) =>
      OxmlElement(qnTagName,
          attrs: {qn('w:w'): w.toString(), qn('w:type'): type});
  static final qnTagName = qn('w:tblW');
  static final qnTcw = qn('w:tcW');

  int get w => getReqAttrVal('w:w', xsdIntConverter);
  set w(int value) => setReqAttrVal('w:w', value, xsdIntConverter);

  String get type => getReqAttrVal('w:type', stTblWidthConverter);
  set type(String value) => setReqAttrVal('w:type', value, stTblWidthConverter);

  Length? get width {
    if (type != 'dxa') return null;
    return Twips(w);
  }

  set width(Length? value) {
    if (value == null) {
      type = 'dxa'; // Default tcW to dxa=0?
      w = 0;
    } else {
      type = 'dxa';
      w = value.twips;
    }
  }
}

/// `<w:vAlign>` element, specifying vertical alignment of cell content.
class CT_VerticalJc extends BaseOxmlElement {
  CT_VerticalJc(super.element);
  static XmlElement create({required WD_CELL_VERTICAL_ALIGNMENT val}) =>
      OxmlElement(qnTagName,
          attrs: {qn('w:val'): wdCellVerticalAlignmentConverter.toXml(val)!});
  static final qnTagName = qn('w:vAlign');

  WD_CELL_VERTICAL_ALIGNMENT get val =>
      getReqAttrVal('w:val', wdCellVerticalAlignmentConverter);
  set val(WD_CELL_VERTICAL_ALIGNMENT value) =>
      setReqAttrVal('w:val', value, wdCellVerticalAlignmentConverter);
}

/// `<w:vMerge>` element, specifying vertical merging behavior of a cell.
class CT_VMerge extends BaseOxmlElement {
  CT_VMerge(super.element);
  static XmlElement create({String? val}) {
    final attrs = <String, String>{};
    if (val == ST_Merge.RESTART) {
      attrs[qn('w:val')] = stMergeConverter.toXml(val)!;
    }
    return OxmlElement(qnTagName, attrs: attrs);
  }

  static final qnTagName = qn('w:vMerge');

  String? get val => getAttrVal('w:val', stMergeConverter);
  set val(String? value) => setAttrVal('w:val', value, stMergeConverter);
}

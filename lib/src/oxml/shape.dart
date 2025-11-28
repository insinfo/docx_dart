/// Path: lib/src/oxml/shape.dart
/// Based on python-docx: docx/oxml/shape.py
/// Custom element classes for shape-related elements like `<w:inline>`.

import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/parser.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart'; // Para Length
import 'package:xml/xml.dart';

/// `<wp:anchor>` element, container for a "floating" shape.
class CT_Anchor extends BaseOxmlElement {
  CT_Anchor(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('wp:anchor');
}

/// `<a:blip>` element, specifies image source and adjustments.
class CT_Blip extends BaseOxmlElement {
  CT_Blip(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('a:blip');

  String? get embed => getAttrVal('r:embed', stRelationshipIdConverter);
  set embed(String? value) =>
      setAttrVal('r:embed', value, stRelationshipIdConverter);

  String? get link => getAttrVal('r:link', stRelationshipIdConverter);
  set link(String? value) =>
      setAttrVal('r:link', value, stRelationshipIdConverter);
}

/// `<pic:blipFill>` element, specifies picture properties like source and stretching.
class CT_BlipFillProperties extends BaseOxmlElement {
  CT_BlipFillProperties(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final blipFill = OxmlElement(qnTagName);
    blipFill.children.add(CT_Blip.create());
    blipFill.children.add(CT_StretchInfoProperties.create());
    return blipFill;
    // --- End Correction ---
  }

  static final qnTagName = qn('pic:blipFill');

  static final _childSequence = [
    CT_Blip.qnTagName,
    qn('a:srcRect'),
    qn('a:tile'),
    CT_StretchInfoProperties.qnTagName
  ];

  CT_Blip? get blip => childOrNull(_childSequence[0]) == null
      ? null
      : CT_Blip(childOrNull(_childSequence[0])!);

  CT_Blip getOrAddBlip() => CT_Blip(getOrAddChild(
      _childSequence[0], _childSequence.sublist(1), CT_Blip.create));

  CT_StretchInfoProperties? get stretch =>
      childOrNull(_childSequence[3]) == null
          ? null
          : CT_StretchInfoProperties(childOrNull(_childSequence[3])!);

  CT_StretchInfoProperties getOrAddStretch() => CT_StretchInfoProperties(
      getOrAddChild(_childSequence[3], [], CT_StretchInfoProperties.create));
}

/// `<a:graphic>` element, container for a DrawingML object (like a picture or chart).
class CT_GraphicalObject extends BaseOxmlElement {
  CT_GraphicalObject(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final graphic = OxmlElement(qnTagName);
    graphic.children
        .add(CT_GraphicalObjectData.create()); // Add default graphicData
    return graphic;
    // --- End Correction ---
  }

  static final qnTagName = qn('a:graphic');

  CT_GraphicalObjectData get graphicData {
    final child = childOrNull(CT_GraphicalObjectData.qnTagName);
    if (child == null)
      throw StateError(
          'Required child <a:graphicData> not found in <a:graphic>');
    return CT_GraphicalObjectData(child);
  }

  CT_GraphicalObjectData getOrAddGraphicData() =>
      CT_GraphicalObjectData(getOrAddChild(CT_GraphicalObjectData.qnTagName, [],
          () => CT_GraphicalObjectData.create()));
}

/// `<a:graphicData>` element ...
class CT_GraphicalObjectData extends BaseOxmlElement {
  CT_GraphicalObjectData(super.element);
  static XmlElement create({String uri = "URI_PLACEHOLDER"}) =>
  OxmlElement(qnTagName, attrs: {'uri': uri});
  static final qnTagName = qn('a:graphicData');

  CT_Picture? get pic => childOrNull(CT_Picture.qnTagName) == null
      ? null
      : CT_Picture(childOrNull(CT_Picture.qnTagName)!);

  String get uri => getReqAttrVal('uri', xsdTokenConverter);
  set uri(String value) => setReqAttrVal('uri', value, xsdTokenConverter);

  void _insertPic(CT_Picture pic) {
    element.children.clear();
    if (pic.element.parent != null) {
      pic.element.parent!.children.remove(pic.element);
    }
    element.children.add(pic.element);
  }
}

/// `<wp:inline>` element ...
class CT_Inline extends BaseOxmlElement {
  CT_Inline(super.element);
  static XmlElement create() => OxmlElement(qnTagName); // Simple create now
  static final qnTagName = qn('wp:inline');

  static final _childSequence = [
    CT_PositiveSize2D.qnExtent,
    CT_PositiveSize2D.qnEffectExtent,
    CT_NonVisualDrawingProps.qnDocPr,
    qn('wp:cNvGraphicFramePr'),
    CT_GraphicalObject.qnTagName,
  ];

  CT_PositiveSize2D get extent {
    final child = childOrNull(CT_PositiveSize2D.qnExtent);
    if (child == null)
      throw StateError('Required child <wp:extent> not found in <wp:inline>');
    return CT_PositiveSize2D(child);
  }

  CT_PositiveSize2D getOrAddExtent() => CT_PositiveSize2D(getOrAddChild(
      _childSequence[0],
      _childSequence.sublist(1),
      () => CT_PositiveSize2D.create(qnTagName: CT_PositiveSize2D.qnExtent)));

  CT_NonVisualDrawingProps get docPr {
    final child = childOrNull(CT_NonVisualDrawingProps.qnDocPr);
    if (child == null)
      throw StateError('Required child <wp:docPr> not found in <wp:inline>');
    return CT_NonVisualDrawingProps(child);
  }

  CT_NonVisualDrawingProps getOrAddDocPr() =>
      CT_NonVisualDrawingProps(getOrAddChild(
          _childSequence[2],
          _childSequence.sublist(3),
          () => CT_NonVisualDrawingProps.create(
              qnTagName: CT_NonVisualDrawingProps.qnDocPr)));

  CT_GraphicalObject get graphic {
    final child = childOrNull(CT_GraphicalObject.qnTagName);
    if (child == null)
      throw StateError('Required child <a:graphic> not found in <wp:inline>');
    return CT_GraphicalObject(child);
  }

  CT_GraphicalObject getOrAddGraphic() => CT_GraphicalObject(
      getOrAddChild(_childSequence[4], [], CT_GraphicalObject.create));

  static CT_Inline newInline(
      Length cx, Length cy, int shapeId, CT_Picture pic) {
    final inlineElement = parseXml(_inlineXmlTemplate());
    final inline = CT_Inline(inlineElement);
    final extent = inline.getOrAddExtent();
    extent.cx = cx;
    extent.cy = cy;
    final docPr = inline.getOrAddDocPr();
    docPr.id = shapeId;
    docPr.name = "Picture $shapeId";
    final graphic = inline.getOrAddGraphic();
    final graphicData = graphic.getOrAddGraphicData();
    graphicData.uri = nsmap['pic']!;
    graphicData._insertPic(pic);
    return inline;
  }

  static CT_Inline newPicInline(
      int shapeId, String rId, String filename, Length cx, Length cy) {
    final picId = 0;
    final pic = CT_Picture.newPicture(picId, filename, rId, cx, cy);
    final inline = CT_Inline.newInline(cx, cy, shapeId, pic);
    return inline;
  }

  static String _inlineXmlTemplate() {
    return '''
      <wp:inline ${nsdecls(['wp', 'a', 'pic', 'r'])}>
        <wp:extent cx="914400" cy="914400"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="0" name="Picture 0"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks xmlns:a="${nsmap['a']}" noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="${nsmap['a']}">
          <a:graphicData uri="${nsmap['pic']}"/>
        </a:graphic>
      </wp:inline>
    '''; // Set pic uri directly in template
  }
}

/// `<wp:docPr>` and `<pic:cNvPr>` element ...
class CT_NonVisualDrawingProps extends BaseOxmlElement {
  CT_NonVisualDrawingProps(super.element);
  static XmlElement create(
          {int id = 0, String name = 'Drawing', String? qnTagName}) =>
      OxmlElement(qnTagName ?? qnDocPr,
        attrs: {'id': id.toString(), 'name': name});
  static final qnDocPr = qn('wp:docPr');
  static final qnCNvPr = qn('pic:cNvPr');

  int get id => getReqAttrVal('id', stDrawingElementIdConverter);
  set id(int value) => setReqAttrVal('id', value, stDrawingElementIdConverter);
  String get name => getReqAttrVal('name', xsdStringConverter);
  set name(String value) => setReqAttrVal('name', value, xsdStringConverter);
}

/// `<pic:cNvPicPr>` element ...
class CT_NonVisualPictureProperties extends BaseOxmlElement {
  CT_NonVisualPictureProperties(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('pic:cNvPicPr');
}

/// `<a:off>` element ...
class CT_Point2D extends BaseOxmlElement {
  CT_Point2D(super.element);
  static XmlElement create({Length? x, Length? y}) {
    final xVal = x ?? Emu(0);
    final yVal = y ?? Emu(0);
    final attrs = <String, String>{
      'x': stCoordinateConverter.toXml(xVal)!,
      'y': stCoordinateConverter.toXml(yVal)!,
    };
    return OxmlElement(qnTagName, attrs: attrs);
  }

  static final qnTagName = qn('a:off');

  Length get x => getReqAttrVal('x', stCoordinateConverter);
  set x(Length value) => setReqAttrVal('x', value, stCoordinateConverter);
  Length get y => getReqAttrVal('y', stCoordinateConverter);
  set y(Length value) => setReqAttrVal('y', value, stCoordinateConverter);
}

/// `<wp:extent>` and `<a:ext>` elements ...
class CT_PositiveSize2D extends BaseOxmlElement {
  CT_PositiveSize2D(super.element);
  static XmlElement create({Length? cx, Length? cy, String? qnTagName}) {
    final xVal = cx ?? Emu(0);
    final yVal = cy ?? Emu(0);
    final attrs = <String, String>{
      'cx': stPositiveCoordinateConverter.toXml(xVal)!,
      'cy': stPositiveCoordinateConverter.toXml(yVal)!,
    };
    return OxmlElement(qnTagName ?? qnExtent, attrs: attrs);
  }

  static final qnExtent = qn('wp:extent');
  static final qnExt = qn('a:ext');
  static final qnEffectExtent = qn('wp:effectExtent');

  Length get cx => getReqAttrVal('cx', stPositiveCoordinateConverter);
  set cx(Length value) =>
      setReqAttrVal('cx', value, stPositiveCoordinateConverter);
  Length get cy => getReqAttrVal('cy', stPositiveCoordinateConverter);
  set cy(Length value) =>
      setReqAttrVal('cy', value, stPositiveCoordinateConverter);
}

/// `<a:prstGeom>` element ...
class CT_PresetGeometry2D extends BaseOxmlElement {
  CT_PresetGeometry2D(super.element);
  static XmlElement create({String prst = 'rect'}) {
    // --- CORRECTION: Create parent then add children ---
    final prstGeom = OxmlElement(qnTagName, attrs: {'prst': prst});
    prstGeom.children.add(OxmlElement(qn('a:avLst'))); // Add empty avLst
    return prstGeom;
    // --- End Correction ---
  }

  static final qnTagName = qn('a:prstGeom');

  String get prst => getReqAttrVal('prst', xsdTokenConverter);
  set prst(String value) => setReqAttrVal('prst', value, xsdTokenConverter);
}

/// `<a:fillRect>` element ...
class CT_RelativeRect extends BaseOxmlElement {
  CT_RelativeRect(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('a:fillRect');
}

/// `<a:stretch>` element ...
class CT_StretchInfoProperties extends BaseOxmlElement {
  CT_StretchInfoProperties(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final stretch = OxmlElement(qnTagName);
    stretch.children.add(CT_RelativeRect.create()); // Add fillRect by default
    return stretch;
    // --- End Correction ---
  }

  static final qnTagName = qn('a:stretch');

  static final _childSequence = [CT_RelativeRect.qnTagName];

  CT_RelativeRect? get fillRect => childOrNull(_childSequence[0]) == null
      ? null
      : CT_RelativeRect(childOrNull(_childSequence[0])!);

  CT_RelativeRect getOrAddFillRect() => CT_RelativeRect(
      getOrAddChild(_childSequence[0], [], CT_RelativeRect.create));
}

/// `<a:xfrm>` element ...
class CT_Transform2D extends BaseOxmlElement {
  CT_Transform2D(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final xfrm = OxmlElement(qnTagName);
    xfrm.children.add(CT_Point2D.create()); // Add default <a:off>
    xfrm.children.add(CT_PositiveSize2D.create(
        qnTagName: CT_PositiveSize2D.qnExt)); // Add default <a:ext>
    return xfrm;
    // --- End Correction ---
  }

  static final qnTagName = qn('a:xfrm');

  static final _childSequence = [CT_Point2D.qnTagName, CT_PositiveSize2D.qnExt];

  CT_Point2D? get off => childOrNull(_childSequence[0]) == null
      ? null
      : CT_Point2D(childOrNull(_childSequence[0])!);
  CT_Point2D getOrAddOff() => CT_Point2D(
      getOrAddChild(_childSequence[0], [_childSequence[1]], CT_Point2D.create));

  CT_PositiveSize2D? get ext => childOrNull(_childSequence[1]) == null
      ? null
      : CT_PositiveSize2D(childOrNull(_childSequence[1])!);
  CT_PositiveSize2D getOrAddExt() => CT_PositiveSize2D(getOrAddChild(
      _childSequence[1],
      [],
      () => CT_PositiveSize2D.create(qnTagName: _childSequence[1])));

  Length? get cx => ext?.cx;
  set cx(Length? value) {
    if (value == null) {
      final extElement = ext;
      if (extElement != null && extElement.cy == Emu(0)) {
        removeChild(CT_PositiveSize2D.qnExt);
      } else if (extElement != null) {
        extElement.element.removeAttribute('cx');
      }
    } else {
      getOrAddExt().cx = value;
    }
  }

  Length? get cy => ext?.cy;
  set cy(Length? value) {
    if (value == null) {
      final extElement = ext;
      if (extElement != null && extElement.cx == Emu(0)) {
        removeChild(CT_PositiveSize2D.qnExt);
      } else if (extElement != null) {
        extElement.element.removeAttribute('cy');
      }
    } else {
      getOrAddExt().cy = value;
    }
  }
}

/// `<pic:spPr>` element ...
class CT_ShapeProperties extends BaseOxmlElement {
  CT_ShapeProperties(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final spPr = OxmlElement(qnTagName);
    spPr.children.add(CT_Transform2D.create()); // Add xfrm by default
    spPr.children.add(CT_PresetGeometry2D.create()); // Add prstGeom by default
    return spPr;
    // --- End Correction ---
  }

  static final qnTagName = qn('pic:spPr');

  static final _childSequence = [
    CT_Transform2D.qnTagName,
    qn('a:custGeom'),
    CT_PresetGeometry2D.qnTagName,
    qn('a:ln'),
    qn('a:effectLst'),
    qn('a:effectDag'),
    qn('a:scene3d'),
    qn('a:sp3d'),
    qn('a:extLst'),
  ];

  CT_Transform2D? get xfrm => childOrNull(_childSequence[0]) == null
      ? null
      : CT_Transform2D(childOrNull(_childSequence[0])!);
  CT_Transform2D getOrAddXfrm() => CT_Transform2D(getOrAddChild(
      _childSequence[0], _childSequence.sublist(1), CT_Transform2D.create));

  CT_PresetGeometry2D? get prstGeom => childOrNull(_childSequence[2]) == null
      ? null
      : CT_PresetGeometry2D(childOrNull(_childSequence[2])!);
  CT_PresetGeometry2D getOrAddPrstGeom() => CT_PresetGeometry2D(getOrAddChild(
      _childSequence[2],
      _childSequence.sublist(3),
      CT_PresetGeometry2D.create));

  Length? get cx => xfrm?.cx;
  set cx(Length? value) {
    getOrAddXfrm().cx = value;
  }

  Length? get cy => xfrm?.cy;
  set cy(Length? value) {
    getOrAddXfrm().cy = value;
  }
}

/// `<pic:nvPicPr>` element ...
class CT_PictureNonVisual extends BaseOxmlElement {
  CT_PictureNonVisual(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final nvPicPr = OxmlElement(qnTagName);
    nvPicPr.children.add(CT_NonVisualDrawingProps.create(
        qnTagName: CT_NonVisualDrawingProps.qnCNvPr));
    nvPicPr.children.add(CT_NonVisualPictureProperties.create());
    return nvPicPr;
    // --- End Correction ---
  }

  static final qnTagName = qn('pic:nvPicPr');

  static final _childSequence = [
    CT_NonVisualDrawingProps.qnCNvPr,
    CT_NonVisualPictureProperties.qnTagName
  ];

  CT_NonVisualDrawingProps get cNvPr {
    final child = childOrNull(_childSequence[0]);
    if (child == null)
      throw StateError('Required child <pic:cNvPr> not found in <pic:nvPicPr>');
    return CT_NonVisualDrawingProps(child);
  }

  CT_NonVisualDrawingProps getOrAddCNvPr() =>
      CT_NonVisualDrawingProps(getOrAddChild(
          _childSequence[0],
          [_childSequence[1]],
          () => CT_NonVisualDrawingProps.create(qnTagName: _childSequence[0])));

  CT_NonVisualPictureProperties get cNvPicPr {
    final child = childOrNull(_childSequence[1]);
    if (child == null)
      throw StateError(
          'Required child <pic:cNvPicPr> not found in <pic:nvPicPr>');
    return CT_NonVisualPictureProperties(child);
  }

  CT_NonVisualPictureProperties getOrAddCNvPicPr() =>
      CT_NonVisualPictureProperties(getOrAddChild(
          _childSequence[1], [], CT_NonVisualPictureProperties.create));
}

/// `<pic:pic>` element ...
class CT_Picture extends BaseOxmlElement {
  CT_Picture(super.element);
  static XmlElement create() {
    // --- CORRECTION: Create parent then add children ---
    final pic = OxmlElement(qnTagName);
    pic.children.add(CT_PictureNonVisual.create());
    pic.children.add(CT_BlipFillProperties.create());
    pic.children.add(CT_ShapeProperties.create());
    return pic;
    // --- End Correction ---
  }

  static final qnTagName = qn('pic:pic');

  static final _childSequence = [
    CT_PictureNonVisual.qnTagName,
    CT_BlipFillProperties.qnTagName,
    CT_ShapeProperties.qnTagName
  ];

  CT_PictureNonVisual get nvPicPr {
    final child = childOrNull(_childSequence[0]);
    if (child == null)
      throw StateError('Required child <pic:nvPicPr> not found in <pic:pic>');
    return CT_PictureNonVisual(child);
  }

  CT_PictureNonVisual getOrAddNvPicPr() => CT_PictureNonVisual(getOrAddChild(
      _childSequence[0],
      _childSequence.sublist(1),
      CT_PictureNonVisual.create));

  CT_BlipFillProperties get blipFill {
    final child = childOrNull(_childSequence[1]);
    if (child == null)
      throw StateError('Required child <pic:blipFill> not found in <pic:pic>');
    return CT_BlipFillProperties(child);
  }

  CT_BlipFillProperties getOrAddBlipFill() =>
      CT_BlipFillProperties(getOrAddChild(_childSequence[1],
          [_childSequence[2]], CT_BlipFillProperties.create));

  CT_ShapeProperties get spPr {
    final child = childOrNull(_childSequence[2]);
    if (child == null)
      throw StateError('Required child <pic:spPr> not found in <pic:pic>');
    return CT_ShapeProperties(child);
  }

  CT_ShapeProperties getOrAddSpPr() => CT_ShapeProperties(
      getOrAddChild(_childSequence[2], [], CT_ShapeProperties.create));

  static CT_Picture newPicture(
      int picId, String filename, String rId, Length cx, Length cy) {
    final picElement =
        CT_Picture.create(); // Creates element with default children
    final pic = CT_Picture(picElement);
    final nvPicPr = pic.getOrAddNvPicPr();
    final cNvPr = nvPicPr.getOrAddCNvPr();
    cNvPr.id = picId;
    cNvPr.name = filename;
    nvPicPr.getOrAddCNvPicPr();
    final blipFill = pic.getOrAddBlipFill();
    blipFill.getOrAddBlip().embed = rId;
    blipFill.getOrAddStretch().getOrAddFillRect();
    final spPr = pic.getOrAddSpPr();
    spPr.getOrAddXfrm();
    spPr.cx = cx;
    spPr.cy = cy;
    spPr.getOrAddPrstGeom();
    return pic;
  }
}

// --- Placeholder Converters ---
final stRelationshipIdConverter = const ST_RelationshipIdConverter();
final xsdTokenConverter = const XsdTokenConverter();
final stDrawingElementIdConverter = const ST_DrawingElementIdConverter();
final xsdStringConverter = const XsdStringConverter();
final stCoordinateConverter = const ST_CoordinateConverter();
final stPositiveCoordinateConverter = const ST_PositiveCoordinateConverter();

class ST_RelationshipIdConverter implements BaseSimpleType<String> {
  const ST_RelationshipIdConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {}
}

class XsdTokenConverter implements BaseSimpleType<String> {
  const XsdTokenConverter();
  @override
  String fromXml(String xmlValue) =>
      xmlValue.trim().replaceAll(RegExp(r'\s+'), ' ');
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {}
}

class ST_DrawingElementIdConverter implements BaseSimpleType<int> {
  const ST_DrawingElementIdConverter();
  @override
  int fromXml(String xmlValue) => int.parse(xmlValue);
  @override
  String? toXml(int? value) => value?.toString();
  @override
  void validate(int value) {
    if (value < 0) throw ArgumentError("DrawingElementId must be non-negative");
  }
}

class XsdStringConverter implements BaseSimpleType<String> {
  const XsdStringConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {}
}

class ST_CoordinateConverter implements BaseSimpleType<Length> {
  const ST_CoordinateConverter();
  @override
  Length fromXml(String xmlValue) => Emu(int.parse(xmlValue));
  @override
  String? toXml(Length? value) => value?.emu.toString();
  @override
  void validate(Length value) {}
}

class ST_PositiveCoordinateConverter implements BaseSimpleType<Length> {
  const ST_PositiveCoordinateConverter();
  @override
  Length fromXml(String xmlValue) => Emu(int.parse(xmlValue));
  @override
  String? toXml(Length? value) => value?.emu.toString();
  @override
  void validate(Length value) {
    if (value.emu < 0)
      throw ArgumentError("PositiveCoordinate must be non-negative");
  }
}

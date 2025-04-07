// docx/oxml/shape.dart
import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/parser.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart'; // Para Length
import 'package:xml/xml.dart';

// Tipos necessários de outros módulos OXML
import 'package:docx_dart/src/oxml/text/font.dart'; // Exemplo, pode não ser necessário diretamente aqui

class CT_Anchor extends BaseOxmlElement { CT_Anchor(super.element); }
class CT_Blip extends BaseOxmlElement {
  CT_Blip(super.element);
  String? get embed;
  set embed(String? value);
  String? get link;
  set link(String? value);
}

class CT_BlipFillProperties extends BaseOxmlElement {
  CT_BlipFillProperties(super.element);
  CT_Blip? get blip;
}
class CT_GraphicalObject extends BaseOxmlElement {
  CT_GraphicalObject(super.element);
  CT_GraphicalObjectData get graphicData;
}
class CT_GraphicalObjectData extends BaseOxmlElement {
  CT_GraphicalObjectData(super.element);
  CT_Picture? get pic;
  String get uri;
  set uri(String value);
  void _insert_pic(CT_Picture pic); // Método privado
}
class CT_NonVisualDrawingProps extends BaseOxmlElement {
  CT_NonVisualDrawingProps(super.element);
  int get id;
  set id(int value);
  String get name;
  set name(String value);
}
class CT_NonVisualPictureProperties extends BaseOxmlElement { CT_NonVisualPictureProperties(super.element); }
class CT_Point2D extends BaseOxmlElement {
 CT_Point2D(super.element);
 int get x; // ST_Coordinate
 set x(int value); // ST_Coordinate
 int get y; // ST_Coordinate
 set y(int value); // ST_Coordinate
}
class CT_PositiveSize2D extends BaseOxmlElement {
  CT_PositiveSize2D(super.element);
  Length get cx;
  set cx(Length value);
  Length get cy;
  set cy(Length value);
}
class CT_PresetGeometry2D extends BaseOxmlElement { CT_PresetGeometry2D(super.element); }
class CT_RelativeRect extends BaseOxmlElement { CT_RelativeRect(super.element); }
class CT_StretchInfoProperties extends BaseOxmlElement { CT_StretchInfoProperties(super.element); }

class CT_Transform2D extends BaseOxmlElement {
  CT_Transform2D(super.element);
  CT_Point2D? get off;
  CT_PositiveSize2D? get ext;

  Length? get cx;
  set cx(Length? value);
  Length? get cy;
  set cy(Length? value);

  CT_PositiveSize2D get_or_add_ext();
}

class CT_ShapeProperties extends BaseOxmlElement {
  CT_ShapeProperties(super.element);
  CT_Transform2D? get xfrm;

  Length? get cx;
  set cx(Length? value);
  Length? get cy;
  set cy(Length? value);

  CT_Transform2D get_or_add_xfrm();
}

class CT_PictureNonVisual extends BaseOxmlElement {
  CT_PictureNonVisual(super.element);
  CT_NonVisualDrawingProps get cNvPr;
}


class CT_Picture extends BaseOxmlElement {
  CT_Picture(super.element);

  CT_PictureNonVisual get nvPicPr;
  CT_BlipFillProperties get blipFill;
  CT_ShapeProperties get spPr;

  static CT_Picture newPicture(int picId, String filename, String rId, Length cx, Length cy) {
     // Implementação
     throw UnimplementedError();
  }
}


class CT_Inline extends BaseOxmlElement {
  CT_Inline(super.element);

  CT_PositiveSize2D get extent;
  CT_NonVisualDrawingProps get docPr;
  CT_GraphicalObject get graphic;

  static CT_Inline newInline(Length cx, Length cy, int shapeId, CT_Picture pic) {
    // Implementação
    throw UnimplementedError();
  }

  static CT_Inline newPicInline(int shapeId, String rId, String filename, Length cx, Length cy) {
    // Implementação
    throw UnimplementedError();
  }
}
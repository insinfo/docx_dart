// docx/oxml/section.dart
import 'package:docx_dart/src/enum/section.dart';
import 'package:docx_dart/src/oxml/shared.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/oxml/table.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:xml/xml.dart';


typedef BlockElement = XmlElement; // Ou uma classe base comum se existir


class CT_HdrFtr extends BaseOxmlElement {
  CT_HdrFtr(super.element);

  List<XmlElement> get inner_content_elements; // Retorna CT_P | CT_Tbl
  List<CT_P> get p_lst;
  List<CT_Tbl> get tbl_lst;

  CT_P add_p();
  void _insert_tbl(CT_Tbl tbl);
}


class CT_HdrFtrRef extends BaseOxmlElement {
  CT_HdrFtrRef(super.element);

  WD_HEADER_FOOTER get type_;
  set type_(WD_HEADER_FOOTER value);

  String get rId;
  set rId(String value);
}


class CT_PageMar extends BaseOxmlElement {
  CT_PageMar(super.element);

  Length? get top;
  set top(Length? value);
  Length? get right;
  set right(Length? value);
  Length? get bottom;
  set bottom(Length? value);
  Length? get left;
  set left(Length? value);
  Length? get header;
  set header(Length? value);
  Length? get footer;
  set footer(Length? value);
  Length? get gutter;
  set gutter(Length? value);
}


class CT_PageSz extends BaseOxmlElement {
  CT_PageSz(super.element);

  Length? get w;
  set w(Length? value);
  Length? get h;
  set h(Length? value);
  WD_ORIENTATION get orient; // Default PORTRAIT
  set orient(WD_ORIENTATION value);
}

class CT_SectType extends BaseOxmlElement {
 CT_SectType(super.element);
 WD_SECTION_START? get val;
 set val(WD_SECTION_START? value);
}

class CT_SectPr extends BaseOxmlElement {
  CT_SectPr(super.element);

  List<CT_HdrFtrRef> get headerReference;
  List<CT_HdrFtrRef> get footerReference;
  CT_SectType? get type;
  CT_PageSz? get pgSz;
  CT_PageMar? get pgMar;
  CT_OnOff? get titlePg;


  CT_HdrFtrRef add_footerReference(WD_HEADER_FOOTER type, String rId);
  CT_HdrFtrRef add_headerReference(WD_HEADER_FOOTER type, String rId);

  Length? get bottom_margin;
  set bottom_margin(Length? value);
  Length? get footer;
  set footer(Length? value);
  CT_HdrFtrRef? get_footerReference(WD_HEADER_FOOTER type);
  CT_HdrFtrRef? get_headerReference(WD_HEADER_FOOTER type);
  Length? get gutter;
  set gutter(Length? value);
  Length? get header;
  set header(Length? value);
  Iterable<BlockElement> iter_inner_content();
  Length? get left_margin;
  set left_margin(Length? value);
  WD_ORIENTATION get orientation;
  set orientation(WD_ORIENTATION value);
  Length? get page_height;
  set page_height(Length? value);
  Length? get page_width;
  set page_width(Length? value);
  CT_SectPr? get preceding_sectPr;
  String remove_footerReference(WD_HEADER_FOOTER type);
  String remove_headerReference(WD_HEADER_FOOTER type);
  Length? get right_margin;
  set right_margin(Length? value);
  WD_SECTION_START get start_type;
  set start_type(WD_SECTION_START? value);
  bool get titlePg_val;
  set titlePg_val(bool? value);
  Length? get top_margin;
  set top_margin(Length? value);

  CT_PageMar get_or_add_pgMar();
  CT_PageSz get_or_add_pgSz();
  CT_OnOff get_or_add_titlePg();
  CT_SectType get_or_add_type();
  void _remove_titlePg();
  void _remove_type();

  CT_SectPr clone(); // Retorna uma cópia profunda
}
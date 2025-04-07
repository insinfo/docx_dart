// docx/oxml/document.dart
// (Definição necessária para blkcntnr.dart)
import 'package:xml/xml.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart'; // Supondo que xmlchemy seja portado/adaptado
import 'package:docx_dart/src/oxml/section.dart'; // Dependência

class CT_Body extends BaseOxmlElement {
  CT_SectPr? get sectPr;
  set sectPr(CT_SectPr? value);

  List<XmlElement> get inner_content_elements;
  List<CT_P> get p_lst;
  List<CT_Tbl> get tbl_lst;

  CT_P add_p();
  void _insert_tbl(CT_Tbl tbl); // Método privado para inserção
  void clear_content();
  CT_SectPr add_section_break();
  CT_SectPr get_or_add_sectPr();
}
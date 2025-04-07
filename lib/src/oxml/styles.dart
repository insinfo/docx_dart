// docx/oxml/styles.dart
import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/oxml/shared.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/oxml/text/parfmt.dart'; // Para CT_PPr
import 'package:docx_dart/src/oxml/text/font.dart'; // Para CT_RPr
import 'package:docx_dart/src/oxml/xmlchemy.dart';

// Função auxiliar
String styleIdFromName(String name) {
  // Implementação
  return name.replaceAll(' ', ''); // Simplificado
}


class CT_LatentStyles extends BaseOxmlElement {
  CT_LatentStyles(super.element);

  List<CT_LsdException> get lsdException_lst;

  int? get count;
  set count(int? value);
  bool? get defLockedState;
  set defLockedState(bool? value);
  bool? get defQFormat;
  set defQFormat(bool? value);
  bool? get defSemiHidden;
  set defSemiHidden(bool? value);
  int? get defUIPriority;
  set defUIPriority(int? value);
  bool? get defUnhideWhenUsed;
  set defUnhideWhenUsed(bool? value);

  bool bool_prop(String attrName);
  CT_LsdException? get_by_name(String name);
  void set_bool_prop(String attrName, bool value);

  // Métodos para adicionar/remover/obter ou adicionar lsdException
  CT_LsdException add_lsdException();
}

class CT_LsdException extends BaseOxmlElement {
  CT_LsdException(super.element);

  bool? get locked;
  set locked(bool? value);
  String get name;
  set name(String value);
  bool? get qFormat;
  set qFormat(bool? value);
  bool? get semiHidden;
  set semiHidden(bool? value);
  int? get uiPriority;
  set uiPriority(int? value);
  bool? get unhideWhenUsed;
  set unhideWhenUsed(bool? value);

  void delete();
  bool? on_off_prop(String attrName);
  void
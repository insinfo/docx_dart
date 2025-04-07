// docx/oxml/numbering.dart
import 'package:docx_dart/src/oxml/parser.dart';
import 'package:docx_dart/src/oxml/shared.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';

class CT_Num extends BaseOxmlElement {
  CT_Num(super.element);

  CT_DecimalNumber get abstractNumId;
  List<CT_NumLvl> get lvlOverride;
  int get numId;
  set numId(int value);

  CT_NumLvl add_lvlOverride(int ilvl);

  static CT_Num newNum(int numId, int abstractNumId) {
     // Implementação
     throw UnimplementedError();
  }
}

class CT_NumLvl extends BaseOxmlElement {
  CT_NumLvl(super.element);

  CT_DecimalNumber? get startOverride;
  int get ilvl;
  set ilvl(int value);

  CT_DecimalNumber add_startOverride(int val);
}

class CT_NumPr extends BaseOxmlElement {
  CT_NumPr(super.element);

  CT_DecimalNumber? get ilvl;
  CT_DecimalNumber? get numId;
  // Setters não são definidos como propriedades diretas na classe base Python original
  // Eles seriam métodos ou setters na classe Dart, se necessários.
  set ilvlVal(int? value); // Exemplo de setter
  set numIdVal(int? value); // Exemplo de setter
}

class CT_Numbering extends BaseOxmlElement {
  CT_Numbering(super.element);

  List<CT_Num> get num_lst;

  CT_Num add_num(int abstractNumId);
  CT_Num num_having_numId(int numId);
  int get _next_numId;
}
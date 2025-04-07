// docx/table.dart
// (Definição necessária para blkcntnr.dart)
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/oxml/table.dart';
import 'package:docx_dart/src/styles/style.dart'; // Para _TableStyle

class Table extends StoryChild {
 final CT_Tbl _tbl;

 Table(this._tbl, ProvidesStoryPart parent) : super(parent);

 set style(Object? style); // Aceita String ou _TableStyle?

 // Definições de propriedades e métodos virão aqui na implementação
}
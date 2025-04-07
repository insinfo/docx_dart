// docx/text/paragraph.dart
// (Definição necessária para blkcntnr.dart)
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';

class Paragraph extends StoryChild {
 final CT_P _p;

 Paragraph(this._p, ProvidesStoryPart parent) : super(parent);

 // Definições de propriedades e métodos virão aqui na implementação
}
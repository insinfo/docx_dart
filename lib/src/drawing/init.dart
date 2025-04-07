// docx/drawing/init.dart
import 'package:docx_dart/src/oxml/drawing.dart'; // Para CT_Drawing
import 'package:docx_dart/src/shared.dart'; // Para Parented, ProvidesStoryPart
import 'package:xml/xml.dart'; // Para XmlElement

class Drawing extends Parented {
  final XmlElement _drawing;
  // _element é herdado/definido em Parented/ElementProxy se aplicável
  // final ProvidesStoryPart _parent; // Herdado de Parented

  Drawing(this._drawing, ProvidesStoryPart parent) : super(parent) {
     // Se ElementProxy for a base de Parented, inicialize _element aqui
     // Exemplo: (this as ElementProxy)._element = _drawing;
  }
}
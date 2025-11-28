// docx/drawing/init.dart
import 'package:docx_dart/src/oxml/drawing.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/types.dart';

class Drawing extends Parented {
  final CT_Drawing drawing;

  Drawing(this.drawing, ProvidesStoryPart parent) : super(parent);
}
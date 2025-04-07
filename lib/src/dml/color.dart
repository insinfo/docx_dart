// docx/dml/color.dart
import 'package:docx_dart/src/enum/dml.dart';
import 'package:docx_dart/src/oxml/simpletypes.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart'; // Supondo BaseOxmlElement
import 'package:xml/xml.dart'; // Para XmlElement

class ColorFormat extends ElementProxy {

  // Supondo que o construtor de ElementProxy aceite o elemento pai
  ColorFormat(BaseOxmlElement rPrParent) : super(rPrParent);

  /// Um valor [RGBColor] ou `null` se nenhuma cor RGB for especificada.
  /// Atribuir um valor [RGBColor] define `type` como [MsoColorType.rgb].
  /// Atribuir `null` remove qualquer especificação de cor.
  RGBColor? get rgb;
  set rgb(RGBColor? value);

  /// Membro de [MsoThemeColorIndex] ou `null` se nenhuma cor de tema for especificada.
  /// Atribuir um membro de [MsoThemeColorIndex] define `type` como [MsoColorType.theme].
  /// Atribuir `null` remove qualquer especificação de cor.
  MsoThemeColorIndex? get themeColor;
  set themeColor(MsoThemeColorIndex? value);

  /// Somente leitura. Membro de [MsoColorType] (RGB, THEME, ou AUTO).
  /// `null` se nenhuma cor for aplicada neste nível.
  MsoColorType? get type;

  /// Retorna `w:rPr/w:color` ou `null` se não presente.
  XmlElement? get _color;
}
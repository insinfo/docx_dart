// docx/blkcntnr.dart
import 'dart:core';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/oxml/section.dart';
import 'package:docx_dart/src/oxml/table.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/table.dart';
import 'package:xml/xml.dart'; // Para XmlElement como substituto de CT_*

// typedef BlockItemElement = XmlElement; // Definido em shared.dart

abstract class BlockItemContainer extends StoryChild {
  final BlockItemElement _element;

  BlockItemContainer(this._element, ProvidesStoryPart parent) : super(parent);

  /// Adiciona um novo parágrafo ao final do conteúdo neste container.
  /// O parágrafo tem `text` em uma única run se presente, e recebe o estilo `style`.
  /// Se `style` for `null`, nenhum estilo de parágrafo é aplicado (efeito 'Normal').
  Paragraph addParagraph({String text = "", ParagraphStyle? style});

  /// Retorna uma tabela de `width` tendo `rows` linhas e `cols` colunas.
  /// A tabela é anexada ao final do conteúdo neste container.
  /// `width` é distribuído igualmente entre as colunas da tabela.
  Table addTable(int rows, int cols, Length width);

  /// Gera cada `Paragraph` ou `Table` neste container na ordem do documento.
  Iterable<dynamic> iterInnerContent(); // Retorna Iterable<Paragraph | Table>

  /// Uma lista contendo os parágrafos neste container, na ordem do documento. Somente leitura.
  List<Paragraph> get paragraphs;

  /// Uma lista contendo as tabelas neste container, na ordem do documento. Somente leitura.
  List<Table> get tables;

  /// Retorna um parágrafo recém-adicionado ao final do conteúdo neste container.
  Paragraph _addParagraph();
}
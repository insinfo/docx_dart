# docx_dart

Port em Dart da biblioteca `python-docx`, com foco em manipular arquivos `.docx` no ecossistema Dart e Flutter sem depender de Python em tempo de execucao.

O projeto ainda nao atingiu paridade total com a biblioteca original, mas ja cobre uma parte relevante do fluxo de criacao e edicao de documentos WordprocessingML: abrir `.docx`, criar documentos a partir do template embutido, editar paragrafos, runs, tabelas, secoes, estilos e inserir imagens inline.

## Estado atual

Esta documentacao foi reescrita a partir de uma comparacao entre:

- `referencias/python-docx/src/docx`
- `lib/src`

O estado atual do port pode ser resumido assim:

- a base OPC de leitura e escrita de pacote `.docx` esta implementada;
- a API principal de documento (`loadDocxDocument()` e `Document`) esta utilizavel;
- paragrafos, runs, hyperlinks, page-breaks e formatacao de paragrafo estao em bom estado;
- tabelas, celulas, merges e alinhamento ja possuem cobertura real de testes;
- secoes e heranca de header/footer estao parcialmente consolidadas;
- o pipeline de imagens inline esta funcionando para PNG, incluindo save e reload;
- o principal gap funcional atual esta em `lib/src/image/`, que ainda suporta apenas PNG.

## O que ja funciona

Com base no codigo e nos testes atuais, a biblioteca ja oferece:

- criacao de documento novo a partir de um template padrao embutido;
- abertura de `.docx` existente a partir de caminho ou bytes;
- salvamento de documento alterado para disco;
- adicao de paragrafos, headings e page breaks;
- adicao e edicao de runs;
- insercao de hyperlinks e manipulacao de partes do texto;
- acesso e aplicacao de estilos;
- criacao e edicao de tabelas;
- merge de celulas e controle de alinhamento e direcao da tabela;
- criacao de secoes e boa parte da logica de header/footer herdado;
- insercao de imagem inline com dimensionamento preservando aspecto;
- leitura e persistencia de imagens inline ja inseridas.

## Limitacoes conhecidas

Os pontos abaixo ainda estao incompletos ou parciais quando comparados com `python-docx`:

- somente PNG esta implementado em `lib/src/image/image.dart`;
- os handlers equivalentes a `image/bmp.py`, `image/gif.py`, `image/jpeg.py` e `image/tiff.py` ainda nao foram portados;
- `lib/src/parts/numbering.dart` ainda contem trecho nao implementado;
- `lib/src/opc/phys_pkg.dart` nao cobre leitura de pacote em formato diretorio;
- existem pontos de `section.dart` apoiados em subclasses internas para header/footer, mas a superficie publica ainda precisa de consolidacao fina;
- partes menos centrais de DrawingML e numeracao ainda nao chegaram ao mesmo nivel do projeto original.

### Compatibilidade web

O projeto agora possui teste real de navegador para abrir, alterar, salvar e recarregar `.docx` inteiramente em memoria. Isso elimina a dependencia de `dart:io` no fluxo principal da biblioteca para uso em aplicacoes web.

Ao mesmo tempo, existe uma restricao importante do backend JavaScript usado no navegador: inteiros Dart acima da faixa de `Number.MAX_SAFE_INTEGER` nao podem ser representados com seguranca. Por isso, validadores baseados em `xsd:long` e `xsd:unsignedLong` foram ajustados em `lib/src/oxml/simpletypes.dart` para trabalhar com a faixa segura no browser:

```dart
class XsdLongConverter extends _BaseIntConverter {
  // On the JS backend used by browser tests, ints are limited to the safe
  // integer range.
  static const int minInclusive = -9007199254740991;
  static const int maxInclusive = 9007199254740991;

  @override
  void validate(int value) {
    // No range check needed for standard 64-bit int in Dart
    _validateInt(value);
  }
}
```

Na pratica, isso significa que a biblioteca compila e passa nos testes web atuais, mas qualquer ponto futuro que dependa de valores inteiros alem dessa faixa precisara de tratamento especifico para o backend JS.

## Exemplo rapido

```dart
import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/shared.dart';

void main() {
  final document = docx.loadDocxDocument();

  document.addHeading(text: 'Relatorio', level: 1);
  document.addParagraph(text: 'Documento gerado em Dart.');

  final table = document.addTable(2, 2, style: 'Table Grid');
  table.cell(0, 0).text = 'A1';
  table.cell(0, 1).text = 'B1';

  document.addPicture('caminho/para/imagem.png', width: Inches(2));

  document.save('saida.docx');
}
```

## Superficie publica principal

Hoje a exportacao principal em `lib/docx_dart.dart` e enxuta:

- `loadDocxDocument()` para abrir ou criar um documento;
- `Document` como objeto principal de manipulacao.

No estado atual, a API mais estavel gira em torno destes metodos e colecoes do `Document`:

- `addHeading()`
- `addParagraph()`
- `addPageBreak()`
- `addPicture()`
- `addSection()`
- `addTable()`
- `save()`
- `paragraphs`
- `tables`
- `sections`
- `styles`
- `settings`
- `inlineShapes`

## Estrutura do port

O repositorio segue de perto a organizacao do projeto original:

- `lib/src/opc`: infraestrutura de pacote Open Packaging Convention;
- `lib/src/oxml`: wrappers e helpers para o XML OOXML;
- `lib/src/text`: API de alto nivel para paragrafo, run e formatacao;
- `lib/src/parts`: partes do documento (`DocumentPart`, `HeaderPart`, `ImagePart`, etc.);
- `lib/src/styles`: estilos e resolucao de estilos;
- `lib/src/image`: metadados e suporte a formatos de imagem;
- `lib/src/table.dart`, `section.dart`, `shape.dart` e `document.dart`: APIs principais de alto nivel.

## Validacao atual

O repositorio possui testes Dart cobrindo o que hoje e mais importante no port:

- `test/document_test.dart`
- `test/section_test.dart`
- `test/table_test.dart`
- `test/oxml_constructors_test.dart`
- `test/package_image_test.dart`
- `test/image_png_test.dart`
- `test/browser_document_test.dart`
- `test/base_header_footer_test.dart`

Os fixtures necessarios para esses testes ja foram internalizados em `test/test_files`, entao o projeto nao depende da arvore de referencia para executar a suite atual.

## CI

O repositorio possui workflow de GitHub Actions para validar automaticamente:

- resolucao de dependencias com `dart pub get`;
- analise estatica com `dart analyze`;
- testes Dart em VM com `dart test`;
- teste browser em Chrome com `dart test -p chrome test/browser_document_test.dart`.

## O que ainda falta portar

Comparando com `referencias/python-docx/src/docx`, as proximas lacunas mais relevantes sao:

1. completar `lib/src/image/` com BMP, GIF, JPEG e TIFF;
2. fechar o que falta em numeracao e listas;
3. refinar a superficie de header/footer e os pontos restantes de secao;
4. expandir o suporte a DrawingML alem do fluxo atual de imagem inline.

## Referencia usada no port

O codigo de referencia continua versionado em `referencias/python-docx`, mas a biblioteca Dart nao deve depender dele em tempo de execucao. Ele serve como fonte de comparacao, consulta e orientacao para as proximas etapas do port.

# TODO
continue portando o C:\MyDartProjects\docx_dart\python-docx para dart C:\MyDartProjects\docx_dart\lib\src e atualizando o TODO.md

continue a portar os testes de C:\MyDartProjects\docx_dart\python-docx\tests
e atualizando o TODO.md
## Progresso recente
- [x] Portar `docx/section.py` para `lib/src/section.dart`, incluindo gerenciamento de headers/footers.
- [x] Portar `docx/shape.py` para `lib/src/shape.dart`, expondo `InlineShapes` e `InlineShape`.
- [x] Portar `docx/text/run.py` e utilitários associados (`Hyperlink`, `ParagraphFormat`, `TabStops`, `RenderedPageBreak`).
- [x] Implementar `lib/src/text/paragraph.dart`, cobrindo `addRun`, hyperlinks, page-breaks e formatação.
- [x] Portar `docx/table.py` para `lib/src/table.dart`, incluindo `_Cell`, `_Row`, `_Column`, coleções, e conectar com `BlockItemContainer`/`CT_Tbl` helpers.
- [x] Implementar herança real de headers/footers refinando `CT_SectPr.precedingSectPr` e `iterInnerContent()`.
- [x] Embutir `default.docx` e XML padrões direto no pacote Dart e remover os arquivos originais de `lib/src/templates` e `python-docx/src/docx/templates`.
- [x] Portar casos básicos de `python-docx/tests/test_section.py` (tamanhos de página, margens e orientação) para `test/section_test.dart`.
- [x] Cobrir `Sections` (indexação) e cenários de headers/footers linkados/desvinculados inspirados no restante de `test_section.py`.
- [x] Exercitar `startType` no nível de `CT_SectPr` e adicionar `Sections.slice` para paridade com o slicing de Python.
- [x] Portar cenários de `_BaseHeaderFooter` de `test_section.py` via `BaseHeaderFooterHarness`, incluindo `_getOrAddDefinition` e herança de headers/footers.
- [x] Acrescentar cobertura de merges, alinhamento e `table_direction` em `Table` para paridade com `test_table.py`.
- [x] Validar integração das seções com documentos contendo imagens, garantindo que headers continuem configuráveis.
- [x] Exercitar herança de headers/footers em cadeia para validar o novo cálculo de blocos.
- [x] Adicionar detecção completa de PNG via package:image para expor dimensões/dpi corretos e ajustar o cálculo de escala.

## Próximos passos
- [ ] Conectar `Image` ao pipeline de desenho implementando `OxmlConstructors.newPicInline` e validando inserção real de figuras.
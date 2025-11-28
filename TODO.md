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

## Próximos passos
- [ ] Acrescentar cobertura de testes/manuais para mesclas de células, alinhamento e `table_direction` após o novo proxy.
- [ ] Exercitar herança de headers/footers entre seções (link/unlink) para validar o novo cálculo de blocos.
- [ ] Validar integração dos novos módulos com operações de documento (criação de seções, inserção de imagens).
- [ ] Portar os cenários restantes de `test_section.py` (headers/footers dedicados, tipos de início e tamanhos de página específicos) e cobrir `Sections` slicing/len.

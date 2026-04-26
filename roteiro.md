# Roteiro do Port `python-docx` -> `docx_dart`

Este documento registra o estado tecnico do port com base em duas fontes:

- referencia original em `referencias/python-docx/src/docx`
- implementacao atual em `lib/src`

O objetivo aqui e orientar o trabalho restante com prioridade clara, separando o que ja esta suficientemente portado do que ainda esta parcial ou ausente.

## Resumo executivo

O port ja cobre o nucleo de manipulacao de documentos `.docx`:

- pacote OPC;
- leitura e escrita de documento;
- paragrafos e runs;
- hyperlinks e page-breaks;
- tabelas;
- estilos;
- secoes;
- imagens inline PNG.

Em termos praticos, o projeto esta num ponto em que ja serve para criacao e edicao basica de documentos WordprocessingML. O maior gap funcional atual e a ausencia de suporte completo a formatos de imagem alem de PNG.

## Blocos considerados consolidados

Os itens abaixo ja possuem implementacao suficientemente avancada para deixar de ser prioridade imediata.

### Infraestrutura OPC

Arquivos Dart:

- `lib/src/opc/constants.dart`
- `lib/src/opc/packuri.dart`
- `lib/src/opc/part.dart`
- `lib/src/opc/package.dart`
- `lib/src/opc/pkgreader.dart`
- `lib/src/opc/pkgwriter.dart`
- `lib/src/opc/rel.dart`
- `lib/src/opc/oxml.dart`

Correspondencia principal:

- `referencias/python-docx/src/docx/opc/*`

Leitura:

- a base de leitura e escrita de `.docx` ja sustenta abertura, alteracao e persistencia de documentos, inclusive com imagens inline.

### API principal de documento

Arquivos Dart:

- `lib/src/api.dart`
- `lib/src/document.dart`
- `lib/src/package.dart`
- `lib/src/parts/document.dart`

Correspondencia principal:

- `referencias/python-docx/src/docx/api.py`
- `referencias/python-docx/src/docx/document.py`
- `referencias/python-docx/src/docx/package.py`
- `referencias/python-docx/src/docx/parts/document.py`

Estado:

- criar, abrir e salvar documentos ja funciona;
- `addParagraph`, `addHeading`, `addTable`, `addSection` e `addPicture` ja existem.

### Texto, paragrafos e runs

Arquivos Dart:

- `lib/src/text/paragraph.dart`
- `lib/src/text/run.dart`
- `lib/src/text/parfmt.dart`
- `lib/src/text/tabstops.dart`
- `lib/src/text/hyperlink.dart`
- `lib/src/text/pagebreak.dart`
- `lib/src/oxml/text/*`

Correspondencia principal:

- `referencias/python-docx/src/docx/text/*`
- `referencias/python-docx/src/docx/oxml/text/*`

Estado:

- este e um dos blocos mais maduros do port.

### Tabelas

Arquivos Dart:

- `lib/src/table.dart`
- `lib/src/oxml/table.dart`

Correspondencia principal:

- `referencias/python-docx/src/docx/table.py`
- `referencias/python-docx/src/docx/oxml/table.py`

Estado:

- criacao de tabelas, acesso a linhas, colunas, celulas, merges e propriedades principais ja estao implementados.

### Estilos

Arquivos Dart:

- `lib/src/styles/*`
- `lib/src/parts/styles.dart`
- `lib/src/oxml/styles.dart`

Correspondencia principal:

- `referencias/python-docx/src/docx/styles/*`
- `referencias/python-docx/src/docx/parts/styles.py`
- `referencias/python-docx/src/docx/oxml/styles.py`

Estado:

- o suporte ja e suficiente para os fluxos mais comuns de resolucao e aplicacao de estilo.

## Blocos parcialmente portados

### Imagens

Arquivos Dart:

- `lib/src/image/image.dart`
- `lib/src/image/helpers.dart`
- `lib/src/parts/image.dart`
- `lib/src/shape.dart`
- `lib/src/oxml/shape.dart`
- `lib/src/oxml/oxml_constructors.dart`

Referencia original:

- `referencias/python-docx/src/docx/image/image.py`
- `referencias/python-docx/src/docx/image/png.py`
- `referencias/python-docx/src/docx/image/jpeg.py`
- `referencias/python-docx/src/docx/image/gif.py`
- `referencias/python-docx/src/docx/image/bmp.py`
- `referencias/python-docx/src/docx/image/tiff.py`

Estado real:

- o fluxo de `addPicture()` com PNG esta funcionando;
- dimensionamento e round-trip de imagem inline ja possuem testes;
- deduplicacao por `sha1` e escrita de `ImagePart` ja funcionam;
- ainda falta portar os demais formatos da arvore `image/` da referencia.

Impacto:

- esta e a maior lacuna funcional visivel para usuario final.

### Secoes e header/footer

Arquivos Dart:

- `lib/src/section.dart`
- `lib/src/parts/hdrftr.dart`
- `lib/src/oxml/section.dart`

Referencia original:

- `referencias/python-docx/src/docx/section.py`
- `referencias/python-docx/src/docx/parts/hdrftr.py`
- `referencias/python-docx/src/docx/oxml/section.py`

Estado real:

- ha boa cobertura da logica de heranca e vinculacao;
- existem pontos internos apoiados em subclasses de `_BaseHeaderFooter` ainda marcados como abstratos no proprio `section.dart`;
- o bloco esta perto de fechar, mas ainda merece uma rodada final de consolidacao de API e revisao das propriedades expostas.

### Numeracao

Arquivos Dart:

- `lib/src/parts/numbering.dart`
- `lib/src/oxml/numbering.dart`

Referencia original:

- `referencias/python-docx/src/docx/parts/numbering.py`
- `referencias/python-docx/src/docx/oxml/numbering.py`

Estado real:

- existe base estrutural;
- `NumberingPart.newPart()` ja cria e relaciona uma parte `numbering.xml` vazia;
- listas e fluxos de numeracao de paragrafos ainda nao podem ser tratados como concluidos.

### DrawingML

Arquivos Dart:

- `lib/src/drawing/init.dart`
- `lib/src/oxml/drawing.dart`
- `lib/src/oxml/shape.dart`

Referencia original:

- `referencias/python-docx/src/docx/drawing/*`
- `referencias/python-docx/src/docx/shape.py`

Estado real:

- o necessario para imagem inline ja existe;
- o restante de DrawingML ainda esta mais para infraestrutura do que para superficie completa.

## Pontos ainda ausentes ou claramente incompletos

### Formatos de imagem alem de PNG

Falta portar a logica equivalente a:

- `referencias/python-docx/src/docx/image/bmp.py`
- `referencias/python-docx/src/docx/image/gif.py`
- `referencias/python-docx/src/docx/image/jpeg.py`
- `referencias/python-docx/src/docx/image/tiff.py`

### Suporte completo de numeracao

Parcialmente concluido:

- `lib/src/parts/numbering.dart`

Ainda falta:

- API de alto nivel para listas e numeracao de paragrafos;
- leitura e criacao consistente de definicoes concretas e abstratas de numeracao.

### Leitura de pacote por diretorio

Arquivo com lacuna explicita:

- `lib/src/opc/phys_pkg.dart`

Observacao:

- nao e prioridade alta para a biblioteca em uso normal, mas permanece como diferenca frente a uma infraestrutura mais completa.

## Ordem sugerida de trabalho

### Prioridade 1

Completar `lib/src/image/`.

Motivo:

- maior impacto para usuario final;
- referencia original e bem delimitada;
- ja existe pipeline de insercao pronto, falta ampliar o parser e os metadata handlers.

Entrega esperada:

- suporte a JPEG, GIF, BMP e TIFF;
- testes equivalentes aos casos principais de `python-docx/tests/image/test_image.py`.

### Prioridade 2

Fechar numeracao.

Motivo:

- listas e numeracao sao recurso central do dominio Word;
- ainda ha parte explicitamente nao implementada.

Entrega esperada:

- API de alto nivel para aplicar listas a paragrafos;
- criacao e leitura consistente de definicoes de numeracao.

### Prioridade 3

Consolidar definitivamente header/footer e secao.

Motivo:

- o bloco ja esta avancado;
- restam pontos mais de acabamento e robustez do que de estrutura.

Entrega esperada:

- revisar superficie publica;
- remover arestas internas que ainda parecem abstratas ou provisórias.

### Prioridade 4

Expandir DrawingML alem de imagem inline.

Motivo:

- menor urgencia que imagem e numeracao;
- depende de mais massa de XML especializado.

Entrega esperada:

- mapear o subconjunto de maior valor vindo de `drawing/` e `shape.py`.

## Guia pratico para a proxima rodada

Se a proxima etapa for continuar o port, a sequencia mais eficiente e:

1. portar `referencias/python-docx/src/docx/image/jpeg.py`;
2. portar `referencias/python-docx/src/docx/image/gif.py`;
3. portar `referencias/python-docx/src/docx/image/bmp.py`;
4. portar `referencias/python-docx/src/docx/image/tiff.py`;
5. portar testes de `referencias/python-docx/tests/image/test_image.py` para Dart;
6. concluir `lib/src/parts/numbering.dart`.

## Criterio de conclusao do port

O port pode ser considerado funcionalmente maduro quando os seguintes pontos forem verdadeiros:

- `addPicture()` aceitar os formatos principais suportados pelo `python-docx`;
- numeracao e listas estiverem completas;
- secao e header/footer nao tiverem mais arestas provisórias relevantes;
- o roteiro deixar de ser dominado por infraestrutura faltante e passar a tratar apenas expansoes opcionais.

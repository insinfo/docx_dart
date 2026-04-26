# docx_dart

[![CI](https://github.com/insinfo/docx_dart/actions/workflows/ci.yml/badge.svg)](https://github.com/insinfo/docx_dart/actions/workflows/ci.yml)

`docx_dart` is a Dart library for creating, reading, editing, and saving Microsoft Word `.docx` files. It works with WordprocessingML packages directly, so documents can be generated or modified from Dart and Flutter applications without requiring Microsoft Word or a Python runtime.

This library is a Dart port of the Python `python-docx` project, adapted for the Dart and Flutter ecosystem.

## Features

- Create a new document from the built-in default template.
- Open existing `.docx` files from a path, bytes, or another supported package source.
- Save modified documents back to disk or memory-backed package output.
- Add paragraphs, headings, page breaks, runs, hyperlinks, and basic text formatting.
- Create and edit tables, rows, columns, cells, merged cells, alignment, and table direction.
- Work with sections, page setup, headers, footers, first-page headers/footers, and even-page headers/footers.
- Insert inline images with aspect-ratio-preserving sizing.
- Add Word fields such as table of contents (`TOC`) and page numbers (`PAGE`).
- Configure Word to update document fields when the file is opened.
- Restart page numbering in a section.
- Use the package in browser-oriented flows that operate on document bytes in memory.

## Installation

Add the package to your `pubspec.yaml`:

```yaml
dependencies:
  docx_dart: ^1.0.0
```

Then install dependencies:

```sh
dart pub get
```

## Quick Start

```dart
import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/shared.dart';

void main() {
  final document = docx.loadDocxDocument();

  document.addHeading(text: 'Quarterly Report', level: 1);
  document.addParagraph(text: 'This document was generated with Dart.');

  final table = document.addTable(2, 2, style: 'Table Grid');
  table.cell(0, 0).text = 'Metric';
  table.cell(0, 1).text = 'Value';
  table.cell(1, 0).text = 'Revenue';
  table.cell(1, 1).text = r'$42,000';

  document.addPicture('assets/chart.png', width: Inches(3));

  document.save('report.docx');
}
```

## Opening Documents

Create a blank document from the embedded default template:

```dart
final document = docx.loadDocxDocument();
```

Open an existing document from a file path:

```dart
final document = docx.loadDocxDocument('template.docx');
```

Open a document from bytes:

```dart
final document = docx.loadDocxDocument(bytes);
```

## Paragraphs And Runs

```dart
final paragraph = document.addParagraph(text: 'Hello from Dart.');
paragraph.addRun(' Additional text in the same paragraph.');
```

Headings can be added with Word heading levels:

```dart
document.addHeading(text: 'Introduction', level: 1);
document.addHeading(text: 'Background', level: 2);
```

## Tables

```dart
final table = document.addTable(3, 3, style: 'Table Grid');

table.cell(0, 0).text = 'Name';
table.cell(0, 1).text = 'Department';
table.cell(0, 2).text = 'Status';

table.cell(1, 0).text = 'Ana';
table.cell(1, 1).text = 'Finance';
table.cell(1, 2).text = 'Approved';
```

Cells can contain paragraphs and nested tables:

```dart
final cell = table.cell(2, 0);
cell.text = 'Details';
cell.addParagraph().addRun('Additional notes');
```

## Images

```dart
import 'package:docx_dart/src/shared.dart';

document.addPicture('assets/logo.png', width: Inches(1.5));
```

When only one dimension is provided, the other dimension is calculated from the image metadata to preserve the original aspect ratio.

## Sections, Headers, And Footers

```dart
final section = document.addSection();

section.header.isLinkedToPrevious = false;
section.footer.isLinkedToPrevious = false;

section.header.paragraphs.first.text = 'Confidential Report';

final footerParagraph = section.footer.paragraphs.first.clear();
footerParagraph.addRun('Page ');
footerParagraph.addRun().addPageNumber();
```

Enable different first-page or odd/even headers and footers when needed:

```dart
section.differentFirstPageHeaderFooter = true;
document.settings.oddAndEvenPagesHeaderFooter = true;
```

## Page Numbering

Start page numbering at a specific value for a section:

```dart
final contentSection = document.addSection();
contentSection.pageNumberStart = 1;
```

This is useful when a document has a cover page or preliminary pages before the main content.

## Table Of Contents

`docx_dart` can insert a Word table-of-contents field. Word calculates the visible TOC entries when fields are updated.

```dart
document.addTableOfContents(
  minHeadingLevel: 1,
  maxHeadingLevel: 3,
  customStyleLevels: const {'Custom Title': 1},
);
```

The document is marked to update fields on open:

```dart
document.settings.updateFieldsOnOpen = true;
```

## Cover Page, TOC, And Numbered Content

```dart
final document = docx.loadDocxDocument();

document.addParagraph(text: 'Annual Report', style: 'Title');
document.addPageBreak();

document.addTableOfContents(minHeadingLevel: 1, maxHeadingLevel: 3);
document.addPageBreak();

final contentSection = document.addSection();
contentSection.header.isLinkedToPrevious = false;
contentSection.footer.isLinkedToPrevious = false;
contentSection.pageNumberStart = 1;

contentSection.header.paragraphs.first.text = 'Annual Report';

final footer = contentSection.footer.paragraphs.first.clear();
footer.addRun('Page ');
footer.addRun().addPageNumber();

document.addHeading(text: 'Executive Summary', level: 1);
document.addParagraph(text: 'Main document content starts here.');

document.save('annual-report.docx');
```

## Public API

The main library export is intentionally small:

```dart
import 'package:docx_dart/docx_dart.dart';
```

The primary entry point is:

- `loadDocxDocument()`

The central document object provides these commonly used members:

- `addHeading()`
- `addParagraph()`
- `addPageBreak()`
- `addPicture()`
- `addTableOfContents()`
- `addSection()`
- `addTable()`
- `save()`
- `paragraphs`
- `tables`
- `sections`
- `styles`
- `settings`
- `inlineShapes`

## Testing

Run the full test suite with:

```sh
dart test
```

Run static analysis with:

```sh
dart analyze
```

Run the browser document test with Chrome:

```sh
dart test -p chrome test/browser_document_test.dart
```

## CI

The GitHub Actions workflow validates the package on Ubuntu 24.04 with Dart 3.6.2. It installs dependencies, runs static analysis, executes VM tests, and runs the browser document test in Chrome.

## Acknowledgements

`docx_dart` is a Dart port of [`python-docx`](https://github.com/python-openxml/python-docx). The original project is licensed under the MIT License and is copyright (c) 2013 Steve Canny.

See [NOTICE](NOTICE) for the upstream attribution and license notice.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

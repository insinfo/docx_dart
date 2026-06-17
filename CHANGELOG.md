# Changelog

## 1.2.0

- Added `InlineParagraphStyle` and `InlineRunStyle` for simplified inline styling.
- Added `DocxUnit` helper class for easier dimension definitions (`cm` and `pt`).
- Added `run.addFloatingPicture` to support anchoring images behind text (`<wp:anchor behindDoc="1">`).
- Exposed simplified enums `ParagraphAlignment` and `Underline` for inline styles.
- Exposed internal formatting elements (`InlineParagraphStyle`, `InlineRunStyle`, `ParagraphAlignment`, `Underline`, `DocxUnit`) and `Font` in the barrel file `docx_dart.dart`.
- Added `font` and `paragraphFormat` properties to `ParagraphStyle` and `CharacterStyle` to enable reading and updating formatting/font details directly from styles.
- Added `clear()` and `removeParagraph(Paragraph paragraph)` methods to `BlockItemContainer` to allow removing content without raw XML manipulation.

## 1.1.0

- Relaxed the `image` dependency constraint to support projects that resolve `archive` 3.x through other packages.

## 1.0.0

- Initial public release of `docx_dart`.
- Added support for creating, opening, editing, and saving `.docx` documents.
- Added document APIs for paragraphs, headings, tables, sections, headers, footers, inline images, table of contents fields, and page number fields.
- Added test coverage for document structure, sections, tables, image handling, settings, numbering parts, and browser document workflows.

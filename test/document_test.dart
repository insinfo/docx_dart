import 'dart:io';

import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:test/test.dart';

void main() {
  docx.Document newDocument() => docx.loadDocxDocument();

  group('Document', () {
    group('addHeading', () {
      const headingStyles = {
        0: 'Title',
        1: 'Heading 1',
        2: 'Heading 2',
        9: 'Heading 9',
      };

      headingStyles.forEach((level, expectedStyle) {
        test('sets "$expectedStyle" for level $level', () {
          final document = newDocument();
          final paragraph =
              document.addHeading(text: 'Spam vs. Bacon', level: level);

          expect(paragraph.text, 'Spam vs. Bacon');
          expect(paragraph.style?.name, expectedStyle);

          final lastParagraph = document.paragraphs.last;
          expect(lastParagraph.text, 'Spam vs. Bacon');
          expect(lastParagraph.style?.name, expectedStyle);
        });
      });

      test('rejects heading levels outside 0-9', () {
        final document = newDocument();

        void expectLevelError(int level) {
          expect(
            () => document.addHeading(level: level),
            throwsA(
              isA<ArgumentError>().having(
                (error) => error.message,
                'message',
                contains('level must be in range 0-9'),
              ),
            ),
          );
        }

        expectLevelError(-1);
        expectLevelError(10);
      });
    });

    test('addParagraph appends text with provided style', () {
      final document = newDocument();
      final initialCount = document.paragraphs.length;

      final paragraph =
          document.addParagraph(text: 'Hello, Paragraph', style: 'Heading 1');

      expect(paragraph.text, 'Hello, Paragraph');
      expect(paragraph.style?.name, 'Heading 1');

      final paragraphs = document.paragraphs;
      expect(paragraphs.length, initialCount + 1);
      expect(paragraphs.last.text, 'Hello, Paragraph');
      expect(paragraphs.last.style?.name, 'Heading 1');
    });

    test('addTable creates table with requested shape and style', () {
      final document = newDocument();
      final initialTableCount = document.tables.length;

      final table = document.addTable(2, 3, style: 'Table Grid');

      expect(table.rows.length, 2);
      expect(table.columns.length, 3);
      expect(table.style?.name, 'Table Grid');

      final tables = document.tables;
      expect(tables.length, initialTableCount + 1);
      final insertedTable = tables.last;
      expect(insertedTable.rows.length, 2);
      expect(insertedTable.columns.length, 3);
      expect(insertedTable.style?.name, 'Table Grid');
    });

    test('documents containing images can add sections with independent headers', () {
      final document = docx.loadDocxDocument(_testFile('having-images.docx'));

      final shapesBefore = document.inlineShapes.length;

      final newSection = document.addSection();
      final newHeader = newSection.header;
      expect(newHeader.isLinkedToPrevious, isTrue);

      newHeader.isLinkedToPrevious = false;
      _ensureParagraph(newHeader).text = 'Image doc header';

      final newHeaderTexts = _paragraphTexts(newHeader);
      expect(newHeaderTexts, contains('Image doc header'));
      expect(document.inlineShapes.length, shapesBefore);
    });
  });
}

Paragraph _ensureParagraph(dynamic container) {
  final paragraphs = (container.paragraphs as List<Paragraph>);
  if (paragraphs.isEmpty) {
    return container.addParagraph();
  }
  return paragraphs.first;
}

List<String> _paragraphTexts(dynamic container) {
  final paragraphs = (container.paragraphs as List<Paragraph>);
  return paragraphs.map((p) => p.text).toList(growable: false);
}

String _testFile(String filename) {
  final relative = 'python-docx/tests/test_files/$filename';
  final file = File(relative);
  if (!file.existsSync()) {
    throw StateError('Expected test file at ${file.path}');
  }
  return file.path;
}

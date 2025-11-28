import 'dart:io';

import 'package:docx_dart/docx_dart.dart';
import 'package:docx_dart/src/enum/section.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/table.dart';
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:test/test.dart';

void main() {
  group('Sections', () {
    test('Document.sections grows when adding a section', () {
      final document = loadDocxDocument();
      final initialCount = document.sections.length;

      expect(initialCount, greaterThanOrEqualTo(1));

      document.addSection();

      expect(document.sections.length, initialCount + 1);
    });

    test('Section.startType can be updated', () {
      final document = loadDocxDocument();
      final section = document.sections.last;

      expect(section.startType, WD_SECTION.NEW_PAGE);

      section.startType = WD_SECTION.NEW_COLUMN;
      expect(section.startType, WD_SECTION.NEW_COLUMN);

      section.startType = null;
      expect(section.startType, WD_SECTION.NEW_PAGE);
    });

    test('Section page size round-trips through setters', () {
      final document = loadDocxDocument();
      final section = document.sections.last;

      section.pageWidth = Inches(5);
      section.pageHeight = Inches(7);

      expect(section.pageWidth, Inches(5));
      expect(section.pageHeight, Inches(7));

      section.pageWidth = null;
      section.pageHeight = null;

      expect(section.pageWidth, isNull);
      expect(section.pageHeight, isNull);
    });

    test('Section margin properties round-trip through setters', () {
      final document = loadDocxDocument();
      final section = document.sections.last;

      section.leftMargin = Inches(1.25);
      section.rightMargin = Inches(1.5);
      section.topMargin = Inches(0.75);
      section.bottomMargin = Inches(1.0);
      section.gutter = Inches(0.2);
      section.headerDistance = Inches(0.5);
      section.footerDistance = Inches(0.6);

      expect(section.leftMargin, Inches(1.25));
      expect(section.rightMargin, Inches(1.5));
      expect(section.topMargin, Inches(0.75));
      expect(section.bottomMargin, Inches(1.0));
      expect(section.gutter, Inches(0.2));
      expect(section.headerDistance, Inches(0.5));
      expect(section.footerDistance, Inches(0.6));

      section.leftMargin = null;
      section.headerDistance = null;
      section.footerDistance = null;

      expect(section.leftMargin, isNull);
      expect(section.headerDistance, isNull);
      expect(section.footerDistance, isNull);
    });

    test('Section orientation toggles between portrait and landscape', () {
      final document = loadDocxDocument();
      final section = document.sections.last;

      expect(section.orientation, WD_ORIENTATION.PORTRAIT);

      section.orientation = WD_ORIENTATION.LANDSCAPE;
      expect(section.orientation, WD_ORIENTATION.LANDSCAPE);

      section.orientation = null;
      expect(section.orientation, WD_ORIENTATION.PORTRAIT);
    });

    test('Section differentFirstPageHeaderFooter toggles flag', () {
      final document = loadDocxDocument();
      final section = document.sections.last;

      expect(section.differentFirstPageHeaderFooter, isFalse);
      section.differentFirstPageHeaderFooter = true;
      expect(section.differentFirstPageHeaderFooter, isTrue);
      section.differentFirstPageHeaderFooter = false;
      expect(section.differentFirstPageHeaderFooter, isFalse);
    });

    test('Section.iterInnerContent yields paragraphs and tables in order', () {
      final path = _testFile('sct-inner-content.docx');
      final document = loadDocxDocument(path);

      expect(document.sections.length, 3);

      final expectations = [
        ['P:P1', 'T:T2', 'P:P3'],
        ['T:T4', 'P:P5', 'P:P6'],
        ['P:P7', 'P:P8', 'P:P9'],
      ];

      for (var index = 0; index < expectations.length; index++) {
        final section = document.sections[index];
        final content = section.iterInnerContent().toList();
        final labels = content.map((item) {
          if (item is Paragraph) {
            return 'P:${item.text}';
          }
          if (item is Table) {
            return 'T:${item.rows.first.cells.first.text}';
          }
          throw StateError('Unexpected content type: ${item.runtimeType}');
        }).toList();

        expect(labels, expectations[index]);
      }
    });
  });
}

String _testFile(String filename) {
  final relative = 'python-docx/tests/test_files/$filename';
  final file = File(relative);
  if (!file.existsSync()) {
    throw StateError('Expected test file at ${file.path}');
  }
  return file.path;
}

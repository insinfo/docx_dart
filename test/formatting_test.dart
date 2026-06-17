import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/shared.dart';
import 'package:test/test.dart';

import 'test_file.dart';

void main() {
  group('DocxUnit', () {
    test('cm returns a Cm instance with correct value', () {
      final unit = docx.DocxUnit.cm(2.5);
      expect(unit, isA<Cm>());
      expect(unit.cm, equals(2.5));
    });

    test('pt returns a Pt instance with correct value', () {
      final unit = docx.DocxUnit.pt(12.0);
      expect(unit, isA<Pt>());
      expect(unit.pt, equals(12.0));
    });
  });

  group('InlineParagraphStyle', () {
    test('applies all style properties correctly to added paragraph', () {
      final doc = docx.loadDocxDocument();
      final style = docx.InlineParagraphStyle(
        alignment: docx.ParagraphAlignment.center,
        spaceBeforePt: 12.0,
        spaceAfterPt: 18.0,
        lineSpacing: 1.5,
        firstLineIndentCm: 1.25,
      );

      final paragraph = doc.addParagraph(text: 'Hello Paragraph', style: style);

      expect(paragraph.alignment, equals(docx.WD_PARAGRAPH_ALIGNMENT.CENTER));
      expect(paragraph.paragraphFormat.spaceBefore?.pt, closeTo(12.0, 0.001));
      expect(paragraph.paragraphFormat.spaceAfter?.pt, closeTo(18.0, 0.001));
      expect(paragraph.paragraphFormat.lineSpacing, equals(1.5));
      expect(paragraph.paragraphFormat.firstLineIndent?.cm, closeTo(1.25, 0.001));
    });

    test('leaves unspecified style properties unmodified', () {
      final doc = docx.loadDocxDocument();
      const style = docx.InlineParagraphStyle();

      final paragraph = doc.addParagraph(text: 'Unstyled Paragraph', style: style);

      expect(paragraph.alignment, isNull);
      expect(paragraph.paragraphFormat.spaceBefore, isNull);
      expect(paragraph.paragraphFormat.spaceAfter, isNull);
      expect(paragraph.paragraphFormat.lineSpacing, isNull);
      expect(paragraph.paragraphFormat.firstLineIndent, isNull);
    });
  });

  group('InlineRunStyle', () {
    test('applies all style properties correctly to added run', () {
      final doc = docx.loadDocxDocument();
      final paragraph = doc.addParagraph();
      const style = docx.InlineRunStyle(
        bold: true,
        italic: true,
        underline: docx.Underline.single,
        fontSizePt: 14.0,
        fontFamily: 'Arial',
      );

      final run = paragraph.addRun('Styled Run', style);

      expect(run.font.bold, isTrue);
      expect(run.font.italic, isTrue);
      expect(run.font.underline, equals(docx.WD_UNDERLINE.SINGLE));
      expect(run.font.size?.pt, closeTo(14.0, 0.001));
      expect(run.font.name, equals('Arial'));
    });

    test('leaves unspecified style properties unmodified', () {
      final doc = docx.loadDocxDocument();
      final paragraph = doc.addParagraph();
      const style = docx.InlineRunStyle();

      final run = paragraph.addRun('Unstyled Run', style);

      expect(run.font.bold, isNull);
      expect(run.font.italic, isNull);
      expect(run.font.underline, isNull);
      expect(run.font.size, isNull);
      expect(run.font.name, isNull);
    });
  });

  group('run.addFloatingPicture', () {
    test('generates an anchor behind document element', () {
      final doc = docx.loadDocxDocument();
      final paragraph = doc.addParagraph();
      final run = paragraph.addRun();
      final pngPath = testFile('300-dpi.png');

      run.addFloatingPicture(pngPath, width: docx.Inches(1.5));

      final xml = paragraph.element.element.toXmlString();
      expect(xml, contains('<wp:anchor'));
      expect(xml, contains('behindDoc="1"'));
      expect(xml, contains('Picture 1'));
      expect(xml, contains('distT="0"'));
    });
  });
}

import 'dart:io';

import 'package:docx_dart/docx_dart.dart';
import 'package:docx_dart/src/enum/section.dart';
import 'package:docx_dart/src/oxml/section.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/table.dart';
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

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

    test('Document.sections reflects added sections when iterated', () {
      final document = loadDocxDocument();
      final initialSections = document.sections.toList(growable: false);

      document.addSection();

      final refreshedSections = document.sections.toList(growable: false);
      expect(refreshedSections.length, document.sections.length);
      expect(refreshedSections.length, initialSections.length + 1);
    });

    test('Sections.slice returns a clamped subset', () {
      final document = loadDocxDocument();
      document.addSection();
      document.addSection();

      final slice = document.sections.slice(1, 3);

      expect(slice.length, 2);
      slice.first.pageWidth = Inches(9);
      expect(document.sections[1].pageWidth, Inches(9));
    });

    test('Sections.slice handles negative indexes', () {
      final document = loadDocxDocument();
      document.addSection();
      document.addSection();

      final lastSection = document.sections.slice(-1).single;
      lastSection.pageHeight = Inches(11);

      expect(document.sections.last.pageHeight, Inches(11));
    });

    test('Sections indexer returns shared section state', () {
      final document = loadDocxDocument();

      document.sections[0].pageWidth = Inches(6.25);

      final fetchedAgain = document.sections[0];
      expect(fetchedAgain.pageWidth, Inches(6.25));
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

    test('Section headers can unlink from previous sections', () {
      final document = loadDocxDocument();
      final initialFirstHeader = document.sections.first.header;
      initialFirstHeader.isLinkedToPrevious = false;
      _ensureParagraph(initialFirstHeader).text = 'First header marker';

      document.addSection();
      final sections = document.sections.toList(growable: false);
      final firstHeader = sections.first.header;
      final secondHeader = sections.last.header;

      expect(secondHeader.isLinkedToPrevious, isTrue);

      secondHeader.isLinkedToPrevious = false;
      expect(secondHeader.isLinkedToPrevious, isFalse);
      _ensureParagraph(secondHeader).text = 'Second header marker';

      final firstTexts = _paragraphTexts(firstHeader);
      final secondTexts = _paragraphTexts(secondHeader);

      expect(firstTexts.contains('Second header marker'), isFalse);
      expect(secondTexts.contains('Second header marker'), isTrue);
    });

    test('Section headers can relink to previous definitions', () {
      final document = loadDocxDocument();
      final firstHeader = document.sections.first.header;
      firstHeader.isLinkedToPrevious = false;
      _ensureParagraph(firstHeader).text = 'Header #1';

      document.addSection();
      final secondHeader = document.sections.last.header;
      secondHeader.isLinkedToPrevious = false;
      _ensureParagraph(secondHeader).text = 'Header #2';

      secondHeader.isLinkedToPrevious = true;

      final updatedTexts = _paragraphTexts(document.sections.last.header);
      expect(updatedTexts, contains('Header #1'));
      expect(updatedTexts, isNot(contains('Header #2')));
    });

    test('Headers cascade across multiple sections until each unlinks', () {
      final document = loadDocxDocument();
      final initialFirstHeader = document.sections.first.header;
      initialFirstHeader.isLinkedToPrevious = false;
      _ensureParagraph(initialFirstHeader).text = 'Header for Section 1';

      document.addSection();
      document.addSection();
      final sections = document.sections.toList(growable: false);
      final firstHeader = sections[0].header;
      final secondHeader = sections[1].header;
      final thirdHeader = sections[2].header;

      expect(secondHeader.isLinkedToPrevious, isTrue);
      expect(thirdHeader.isLinkedToPrevious, isTrue);
      expect(_paragraphTexts(secondHeader), contains('Header for Section 1'));
      expect(_paragraphTexts(thirdHeader), contains('Header for Section 1'));

      secondHeader.isLinkedToPrevious = false;
      _ensureParagraph(secondHeader).text = 'Header for Section 2';

      expect(_paragraphTexts(secondHeader), contains('Header for Section 2'));
      expect(_paragraphTexts(thirdHeader), contains('Header for Section 2'));
      expect(_paragraphTexts(firstHeader), contains('Header for Section 1'));

      thirdHeader.isLinkedToPrevious = false;
      _ensureParagraph(thirdHeader).text = 'Header for Section 3';

      expect(_paragraphTexts(firstHeader), contains('Header for Section 1'));
      expect(_paragraphTexts(secondHeader), contains('Header for Section 2'));
      expect(_paragraphTexts(thirdHeader), contains('Header for Section 3'));
    });

    test('Section even-page headers inherit content until unlinked', () {
      final document = loadDocxDocument();
      final initialEvenHeader = document.sections.first.evenPageHeader;
      initialEvenHeader.isLinkedToPrevious = false;
      _ensureParagraph(initialEvenHeader).text = 'Even header #1';

      document.addSection();
      final sections = document.sections.toList(growable: false);
      final firstEvenHeader = sections.first.evenPageHeader;
      final secondEvenHeader = sections.last.evenPageHeader;

      expect(secondEvenHeader.isLinkedToPrevious, isTrue);

      secondEvenHeader.isLinkedToPrevious = false;
      _ensureParagraph(secondEvenHeader).text = 'Even header #2';

      final firstTexts = _paragraphTexts(firstEvenHeader);
      final secondTexts = _paragraphTexts(secondEvenHeader);

      expect(firstTexts, contains('Even header #1'));
      expect(firstTexts, isNot(contains('Even header #2')));
      expect(secondTexts, contains('Even header #2'));
    });

    test('Section footers can unlink from previous sections', () {
      final document = loadDocxDocument();
      final initialFirstFooter = document.sections.first.footer;
      initialFirstFooter.isLinkedToPrevious = false;
      _ensureParagraph(initialFirstFooter).text = 'First footer marker';

      document.addSection();
      final sections = document.sections.toList(growable: false);
      final firstFooter = sections.first.footer;
      final secondFooter = sections.last.footer;

      expect(secondFooter.isLinkedToPrevious, isTrue);

      secondFooter.isLinkedToPrevious = false;
      expect(secondFooter.isLinkedToPrevious, isFalse);
      _ensureParagraph(secondFooter).text = 'Second footer marker';

      final firstTexts = _paragraphTexts(firstFooter);
      final secondTexts = _paragraphTexts(secondFooter);

      expect(firstTexts.contains('Second footer marker'), isFalse);
      expect(secondTexts.contains('Second footer marker'), isTrue);
    });

    test('Section footers can relink to previous definitions', () {
      final document = loadDocxDocument();
      final firstFooter = document.sections.first.footer;
      firstFooter.isLinkedToPrevious = false;
      _ensureParagraph(firstFooter).text = 'Footer #1';

      document.addSection();
      final secondFooter = document.sections.last.footer;
      secondFooter.isLinkedToPrevious = false;
      _ensureParagraph(secondFooter).text = 'Footer #2';

      secondFooter.isLinkedToPrevious = true;

      final updatedTexts = _paragraphTexts(document.sections.last.footer);
      expect(updatedTexts, contains('Footer #1'));
      expect(updatedTexts, isNot(contains('Footer #2')));
    });

    test('Footers cascade across multiple sections until each unlinks', () {
      final document = loadDocxDocument();
      final initialFirstFooter = document.sections.first.footer;
      initialFirstFooter.isLinkedToPrevious = false;
      _ensureParagraph(initialFirstFooter).text = 'Footer for Section 1';

      document.addSection();
      document.addSection();
      final sections = document.sections.toList(growable: false);
      final firstFooter = sections[0].footer;
      final secondFooter = sections[1].footer;
      final thirdFooter = sections[2].footer;

      expect(secondFooter.isLinkedToPrevious, isTrue);
      expect(thirdFooter.isLinkedToPrevious, isTrue);
      expect(_paragraphTexts(secondFooter), contains('Footer for Section 1'));
      expect(_paragraphTexts(thirdFooter), contains('Footer for Section 1'));

      secondFooter.isLinkedToPrevious = false;
      _ensureParagraph(secondFooter).text = 'Footer for Section 2';

      expect(_paragraphTexts(secondFooter), contains('Footer for Section 2'));
      expect(_paragraphTexts(thirdFooter), contains('Footer for Section 2'));
      expect(_paragraphTexts(firstFooter), contains('Footer for Section 1'));

      thirdFooter.isLinkedToPrevious = false;
      _ensureParagraph(thirdFooter).text = 'Footer for Section 3';

      expect(_paragraphTexts(firstFooter), contains('Footer for Section 1'));
      expect(_paragraphTexts(secondFooter), contains('Footer for Section 2'));
      expect(_paragraphTexts(thirdFooter), contains('Footer for Section 3'));
    });

    test('Section first-page footers honor link-to-previous state', () {
      final document = loadDocxDocument();
      final firstSection = document.sections.first;
      firstSection.differentFirstPageHeaderFooter = true;
      var firstFirstPageFooter = firstSection.firstPageFooter;
      firstFirstPageFooter.isLinkedToPrevious = false;
      _ensureParagraph(firstFirstPageFooter).text = 'First footer #1';

      document.addSection();
      final sections = document.sections.toList(growable: false);
      final firstSectionAfterAdd = sections.first;
      final secondSection = sections.last;
      secondSection.differentFirstPageHeaderFooter = true;

      firstFirstPageFooter = firstSectionAfterAdd.firstPageFooter;
      final secondFirstPageFooter = secondSection.firstPageFooter;

      expect(secondFirstPageFooter.isLinkedToPrevious, isTrue);

      secondFirstPageFooter.isLinkedToPrevious = false;
      _ensureParagraph(secondFirstPageFooter).text = 'First footer #2';

      final firstTexts = _paragraphTexts(firstFirstPageFooter);
      final secondTexts = _paragraphTexts(secondFirstPageFooter);

      expect(firstTexts, contains('First footer #1'));
      expect(firstTexts, isNot(contains('First footer #2')));
      expect(secondTexts, contains('First footer #2'));
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

  group('CT_SectPr start type', () {
    test('reports start type based on underlying XML', () {
      final cases = [
        (_sectPrFromInner(''), WD_SECTION.NEW_PAGE),
        (
          _sectPrFromInner('<w:type />'),
          WD_SECTION.NEW_PAGE,
        ),
        (
          _sectPrFromInner(
              '<w:type w:val="continuous" />'),
          WD_SECTION.CONTINUOUS,
        ),
        (
          _sectPrFromInner(
              '<w:type w:val="oddPage" />'),
          WD_SECTION.ODD_PAGE,
        ),
        (
          _sectPrFromInner(
              '<w:type w:val="evenPage" />'),
          WD_SECTION.EVEN_PAGE,
        ),
        (
          _sectPrFromInner(
              '<w:type w:val="nextColumn" />'),
          WD_SECTION.NEW_COLUMN,
        ),
      ];

      for (final (sectPr, expected) in cases) {
        expect(sectPr.start_type, expected);
      }
    });

    test('writing start type updates XML as python-docx', () {
      final cases = [
        (_sectPrFromInner(''), WD_SECTION.EVEN_PAGE, 'evenPage'),
        (_sectPrFromInner('<w:type />'), WD_SECTION.NEW_COLUMN, 'nextColumn'),
        (_sectPrFromInner('<w:type w:val="oddPage"/>'), null, null),
        (_sectPrFromInner('<w:type w:val="evenPage"/>'), WD_SECTION.NEW_PAGE,
            null),
      ];

      for (final (sectPr, value, expectedVal) in cases) {
        sectPr.start_type = value;
        final xml = XmlDocument.parse(_serializeSectPr(sectPr));
        final typeNodes = xml.rootElement
          .findElements('type', namespace: _wNamespace)
            .toList();
        if (expectedVal == null) {
          expect(typeNodes, isEmpty);
        } else {
          expect(typeNodes, hasLength(1));
          final attr = typeNodes.first.getAttribute('val', namespace: _wNamespace);
          expect(attr, expectedVal);
        }
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

const _wNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

CT_SectPr _sectPrFromInner(String innerXml) {
  final xml = '<w:sectPr xmlns:w="$_wNamespace">$innerXml</w:sectPr>';
  final element = XmlDocument.parse(xml).rootElement;
  return CT_SectPr(element);
}

String _serializeSectPr(CT_SectPr sectPr) =>
    sectPr.element.toXmlString(pretty: false);

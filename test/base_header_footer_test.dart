import 'package:docx_dart/src/enum/section.dart';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/oxml/section.dart';
import 'package:docx_dart/src/package.dart';
import 'package:docx_dart/src/parts/document.dart';
import 'package:docx_dart/src/section.dart';
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('BaseHeaderFooterHarness', () {
    test('knows when it is linked to previous definition', () {
      final linked = _newHarness(hasDefinition: false);
      expect(linked.isLinkedToPrevious, isTrue);

      final unlinked = _newHarness(hasDefinition: true);
      expect(unlinked.isLinkedToPrevious, isFalse);
    });

    test('can change whether it is linked to previous section', () {
      final harness = _newHarness(hasDefinition: true);

      harness.isLinkedToPrevious = true;
      expect(harness.isLinkedToPrevious, isTrue);
      expect(harness.dropDefinitionCalls, 1);

      harness.isLinkedToPrevious = false;
      expect(harness.isLinkedToPrevious, isFalse);
      expect(harness.addDefinitionCalls, 1);
    });

    test('part getter returns the underlying definition', () {
      final harness = _newHarness();

      final part = harness.part;

      expect(part, isNotNull);
      expect(harness.addDefinitionCalls, 1);
    });

    test('paragraph access triggers element resolution', () {
      final harness = _newHarness();
      final paragraphs = harness.paragraphs;
      expect(paragraphs, isA<List>());
    });

    test('returns existing definition without re-adding', () {
      final harness = _newHarness(hasDefinition: true);
      final part = harness.part;

      expect(part, isNotNull);
      expect(harness.addDefinitionCalls, 0);
    });

    test('pulls definition from prior header/footer when linked', () {
      final prior = _newHarness(hasDefinition: true);
      final follower = _newHarness();
      follower.prior = prior;

      final part = follower.part;

      expect(part, same(prior.part));
      expect(follower.addDefinitionCalls, 0);
    });

    test('adds definition when first section and linked', () {
      final harness = _newHarness();

      harness.part;

      expect(harness.addDefinitionCalls, 1);
    });
  });
}

BaseHeaderFooterHarness _newHarness({
  bool hasDefinition = false,
  BaseHeaderFooterHarness? prior,
}) {
  final sectPr = CT_SectPr(CT_SectPr.create());
  final documentPart = _TestDocumentPart();
  final harness = BaseHeaderFooterHarness(
    sectPr,
    documentPart,
    WD_HEADER_FOOTER.PRIMARY,
    hasDefinition: hasDefinition,
  );
  harness.prior = prior;
  return harness;
}

class _TestDocumentPart extends DocumentPart {
  _TestDocumentPart()
      : super(
          PackUri('/word/document.xml'),
          CONTENT_TYPE.WML_DOCUMENT_MAIN,
          _testDocumentElement(),
          Package(),
        );
}

CT_Document _testDocumentElement() {
  const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  final xml = XmlDocument.parse('<w:document xmlns:w="$wNs"><w:body/></w:document>');
  return CT_Document(xml.rootElement);
}

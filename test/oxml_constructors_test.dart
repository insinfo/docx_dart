import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/image/image.dart';
import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/oxml_constructors.dart';
import 'package:docx_dart/src/oxml/shape.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:path/path.dart' as p;
import 'package:test/test.dart';

void main() {
  group('OxmlConstructors.newPicInline', () {
    test('builds an inline picture tree from metadata', () {
      const shapeId = 7;
      const rId = 'rId42';
      const filename = 'diagram.png';
      final cx = Inches(2.25);
      final cy = Inches(1.75);

      final element = OxmlConstructors.newPicInline(
        shapeId: shapeId,
        rId: rId,
        filename: filename,
        cx: cx,
        cy: cy,
      );

      final inline = CT_Inline(element);
      expect(inline.element.name.namespaceUri, equals(nsmap['wp']));
      expect(inline.extent.cx, equals(cx));
      expect(inline.extent.cy, equals(cy));
      expect(inline.docPr.id, equals(shapeId));
      expect(inline.docPr.name, equals('Picture $shapeId'));

      final graphicData = inline.graphic.graphicData;
      expect(graphicData.uri, equals(nsmap['pic']));
      final picture = graphicData.pic;
      expect(picture, isNotNull);
      final pic = picture!;

        expect(pic.nvPicPr.cNvPr.id, equals(0));
        expect(pic.nvPicPr.cNvPr.name, equals(filename));
        expect(pic.blipFill.blip, isNotNull,
          reason: inline.element.toXmlString(pretty: true));
          expect(pic.blipFill.blip!.embed, equals(rId),
            reason: inline.element.toXmlString(pretty: true));
      expect(pic.spPr.cx, equals(cx));
      expect(pic.spPr.cy, equals(cy));
    });
  });

  group('StoryPart.newPicInline', () {
    test('scales requested dimensions using image metadata', () async {
      final document = docx.loadDocxDocument();
      final pngPath =
          p.join('python-docx', 'tests', 'test_files', '300-dpi.png');
      final requestedWidth = Inches(1.25);

      final inlineElement = document.part
          .newPicInline(pngPath, width: requestedWidth);
      final inline = CT_Inline(inlineElement);

      final image = await Image.fromPath(pngPath);
      final (expectedWidth, expectedHeight) =
          image.scaledDimensions(width: requestedWidth);

      expect(inline.extent.cx, equals(expectedWidth));
      expect(inline.extent.cy, equals(expectedHeight));

      final pic = inline.graphic.graphicData.pic;
      expect(pic, isNotNull);
      expect(pic!.spPr.cx, equals(expectedWidth));
      expect(pic.spPr.cy, equals(expectedHeight));
    });
  });
}

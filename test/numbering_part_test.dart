import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/oxml.dart' show parse_xml;
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/package.dart';
import 'package:docx_dart/src/parts/numbering.dart';
import 'package:test/test.dart';

void main() {
  group('NumberingPart', () {
    test('newPart creates an empty numbering part', () {
      final package = Package();

      final numberingPart = NumberingPart.newPart(package);

      expect(numberingPart.partname.uri, '/word/numbering.xml');
      expect(numberingPart.contentType, CONTENT_TYPE.WML_NUMBERING);
      expect(numberingPart.package, same(package));
      expect(numberingPart.numberingDefinitions.length, 0);
      expect(numberingPart.element.element.name.local, 'numbering');
    });

    for (final count in [0, 1, 2, 3]) {
      test('numberingDefinitions counts $count numbering definitions', () {
        final numberingPart = NumberingPart(
          PackUri('/word/numbering.xml'),
          CONTENT_TYPE.WML_NUMBERING,
          parse_xml(_numberingXml(count)),
          Package(),
        );

        expect(numberingPart.numberingDefinitions.length, count);
      });
    }
  });
}

String _numberingXml(int numCount) {
  final nums = List.generate(
    numCount,
    (index) {
      final numId = index + 1;
      return '''<w:num w:numId="$numId"><w:abstractNumId w:val="0"/></w:num>''';
    },
  ).join();

  return '''
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  $nums
</w:numbering>
''';
}

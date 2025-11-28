import 'package:docx_dart/src/image/image.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:path/path.dart' as p;
import 'package:test/test.dart';

void main() {
  group('PNG metadata', () {
    test('Image.fromPath reads physical dimensions from PNG header', () async {
      final pngPath =
          p.join('python-docx', 'tests', 'test_files', '300-dpi.png');
      final image = await Image.fromPath(pngPath);

      expect(image.pxWidth, equals(860));
      expect(image.pxHeight, equals(579));
      expect(image.horzDpi, equals(300));
      expect(image.vertDpi, equals(300));

      final targetWidth = Inches(2);
      final (scaledWidth, scaledHeight) =
          image.scaledDimensions(width: targetWidth);
      expect(scaledWidth.emu, equals(targetWidth.emu));

      final nativeWidth = image.width;
      final nativeHeight = image.height;
      final scale =
          nativeWidth.emu == 0 ? 0.0 : targetWidth.emu / nativeWidth.emu;
      final expectedHeight = nativeHeight * scale;
      expect(scaledHeight.emu, equals(expectedHeight.emu));
    });
  });
}

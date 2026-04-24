import 'package:docx_dart/src/image/image.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:test/test.dart';

import 'test_file.dart';

void main() {
  group('PNG metadata', () {
    test('Image.fromPath reads physical dimensions from PNG header', () async {
      final pngPath = testFile('300-dpi.png');
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

  group('JPEG metadata', () {
    test('Image.fromPath reads default dimensions from baseline JPEG', () async {
      final jpegPath = testFile('python-icon.jpeg');
      final image = await Image.fromPath(jpegPath);

      expect(image.contentType, equals('image/jpeg'));
      expect(image.ext, equals('jpeg'));
      expect(image.pxWidth, equals(204));
      expect(image.pxHeight, equals(204));
      expect(image.horzDpi, equals(72));
      expect(image.vertDpi, equals(72));
    });

    test('Image.fromPath reads physical dimensions from JFIF JPEG header', () async {
      final jpegPath = testFile('300-dpi.jpg');
      final image = await Image.fromPath(jpegPath);

      expect(image.contentType, equals('image/jpeg'));
      expect(image.ext, equals('jpg'));
      expect(image.pxWidth, equals(1504));
      expect(image.pxHeight, equals(1936));
      expect(image.horzDpi, equals(300));
      expect(image.vertDpi, equals(300));

      final targetHeight = Inches(2);
      final (scaledWidth, scaledHeight) =
          image.scaledDimensions(height: targetHeight);

      expect(scaledHeight.emu, equals(targetHeight.emu));

      final nativeWidth = image.width;
      final nativeHeight = image.height;
      final scale =
          nativeHeight.emu == 0 ? 0.0 : targetHeight.emu / nativeHeight.emu;
      final expectedWidth = nativeWidth * scale;
      expect(scaledWidth.emu, equals(expectedWidth.emu));
    });
  });
}

import 'dart:io';

import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/image/image.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/package.dart';
import 'package:docx_dart/src/parts/image.dart';
import 'package:test/test.dart';

import 'test_file.dart';

void main() {
  group('ImagePart', () {
    test('fromImage preserves filename and sha1', () async {
      final image = await Image.fromPath(testFile('monty-truth.png'));

      final imagePart = ImagePart.fromImage(
        image,
        PackUri('/word/media/image1.png'),
      );

      expect(imagePart.filename, equals('monty-truth.png'));
      expect(imagePart.sha1, equals(image.sha1));
    });

    test('load infers filename from partname and decodes image on demand', () {
      final bytes = File(testFile('monty-truth.png')).readAsBytesSync();

      final imagePart = ImagePart.load(
        PackUri('/word/media/image9.png'),
        'image/png',
        bytes,
        Package(),
      );

      expect(imagePart.filename, equals('image.png'));
      expect(imagePart.image.sha1, equals(Image.fromBytes(bytes).sha1));
      expect(imagePart.sha1, equals(Image.fromBytes(bytes).sha1));
    });
  });

  group('Package image pipeline', () {
    test('reuses the same related image part for duplicate image descriptors', () {
      final document = docx.loadDocxDocument();
      final package = document.part.package as Package;
      final imagePath = testFile('300-dpi.png');

      final (firstRid, _) = document.part.getOrAddImage(imagePath);
      final (secondRid, _) = document.part.getOrAddImage(
        File(imagePath).readAsBytesSync(),
      );

      expect(secondRid, equals(firstRid));
      expect(package.parts.whereType<ImagePart>().toList(), hasLength(1));
    });

    test('assigns the next available partname for each image extension', () {
      final document = docx.loadDocxDocument();
      final package = document.part.package as Package;

      document.part.getOrAddImage(testFile('300-dpi.png'));
      document.part.getOrAddImage(testFile('python-icon.jpeg'));

      final imageParts = package.parts.whereType<ImagePart>().toList()
        ..sort((left, right) => left.partname.uri.compareTo(right.partname.uri));

      expect(imageParts, hasLength(2));
      expect(imageParts[0].partname.uri, equals('/word/media/image1.jpeg'));
      expect(imageParts[1].partname.uri, equals('/word/media/image1.png'));
    });

    test('gathers embedded image parts when loading a document with images', () {
      final document = docx.loadDocxDocument(testFile('having-images.docx'));
      final package = document.part.package as Package;
      final imageParts = package.parts.whereType<ImagePart>().toList();

      expect(imageParts, hasLength(3));
      for (final imagePart in imageParts) {
        expect(package.imageParts.contains(imagePart), isTrue);
      }
    });
  });
}


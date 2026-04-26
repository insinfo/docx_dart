import 'dart:typed_data';

import 'package:docx_dart/src/image/exceptions.dart';
import 'package:docx_dart/src/image/helpers.dart';
import 'package:test/test.dart';

void main() {
  group('StreamReader', () {
    test('readString reads a string of specified length at offset', () {
      final streamReader = StreamReader(
        Uint8List.fromList([0x01, 0x02, ...'foobar'.codeUnits, 0x03, 0x04]),
        Endian.big,
      );

      expect(streamReader.readString(6, 2), 'foobar');
    });

    test('readString raises on unexpected EOF', () {
      final streamReader = StreamReader(
        Uint8List.fromList([0x01, 0x02, ...'foobar'.codeUnits, 0x03, 0x04]),
        Endian.big,
      );

      expect(
        () => streamReader.readString(9, 2),
        throwsA(isA<UnexpectedEndOfFileError>()),
      );
    });

    for (final fixture in [
      (Endian.big, [0xBE, 0x00, 0x00, 0x00, 0x2A, 0xEF], 1, 42),
      (Endian.little, [0xBE, 0xEF, 0x2A, 0x00, 0x00, 0x00], 2, 42),
    ]) {
      final (endian, bytes, offset, expectedInt) = fixture;

      test('readLong reads $expectedInt using $endian byte order', () {
        final streamReader = StreamReader(Uint8List.fromList(bytes), endian);

        expect(streamReader.readLong(offset), expectedInt);
      });
    }
  });
}

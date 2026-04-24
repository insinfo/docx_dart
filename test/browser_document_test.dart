@TestOn('browser')

import 'dart:typed_data';

import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:test/test.dart';

void main() {
  group('browser docx support', () {
    test('loads, mutates, saves, and reloads a docx entirely in memory', () {
      final original = docx.loadDocxDocument();

      original.addHeading(text: 'Browser title', level: 1);
      original.addParagraph(text: 'Round-trip through browser memory.');
      original.addTable(2, 2, style: 'Table Grid');

      final sink = BytesBuilder(copy: false);
      original.save(sink);
      final bytes = sink.takeBytes();

      expect(bytes, isA<Uint8List>());
      expect(bytes, isNotEmpty);

      final reloaded = docx.loadDocxDocument(bytes);

      expect(reloaded.paragraphs.any((p) => p.text == 'Browser title'), isTrue);
      expect(
        reloaded.paragraphs.any(
          (p) => p.text == 'Round-trip through browser memory.',
        ),
        isTrue,
      );
      expect(reloaded.tables, hasLength(1));
      expect(reloaded.tables.single.rows, hasLength(2));
      expect(reloaded.tables.single.columns, hasLength(2));
    });

    test('rejects filesystem paths in the browser', () {
      expect(
        () => docx.loadDocxDocument('test/test_files/empty.docx'),
        throwsA(isA<UnsupportedError>()),
      );
    });
  });
}

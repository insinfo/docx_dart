import 'package:docx_dart/docx_dart.dart' as docx;
import 'package:docx_dart/src/enum/table.dart';
import 'package:test/test.dart';

void main() {
  docx.Document newDocument() => docx.loadDocxDocument();

  group('Table formatting', () {
    test('alignment round-trips through the Table API', () {
      final document = newDocument();
      final table = document.addTable(2, 2);

      expect(table.alignment, isNull);

      table.alignment = WD_TABLE_ALIGNMENT.CENTER;
      expect(table.alignment, WD_TABLE_ALIGNMENT.CENTER);

      table.alignment = WD_TABLE_ALIGNMENT.RIGHT;
      expect(table.alignment, WD_TABLE_ALIGNMENT.RIGHT);

      table.alignment = null;
      expect(table.alignment, isNull);
    });

    test('tableDirection toggles between RTL and LTR', () {
      final document = newDocument();
      final table = document.addTable(2, 2);

      expect(table.tableDirection, isNull);

      table.tableDirection = WD_TABLE_DIRECTION.RTL;
      expect(table.tableDirection, WD_TABLE_DIRECTION.RTL);

      table.tableDirection = WD_TABLE_DIRECTION.LTR;
      expect(table.tableDirection, WD_TABLE_DIRECTION.LTR);

      table.tableDirection = null;
      expect(table.tableDirection, isNull);
    });
  });

  group('Cell merging', () {
    test('merging diagonal corners produces a shared rectangular span', () {
      final document = newDocument();
      final table = document.addTable(2, 2);

      final merged = table.cell(0, 0).merge(table.cell(1, 1));

      merged.text = 'merged span';

      expect(merged.gridSpan, 2);
      expect(table.cell(0, 1).text, 'merged span');
      expect(table.cell(1, 0).text, 'merged span');
      expect(table.cell(1, 1).text, 'merged span');
    });

    test('merging vertically stacked cells spans multiple rows', () {
      final document = newDocument();
      final table = document.addTable(3, 1);

      final merged = table.cell(0, 0).merge(table.cell(2, 0));

      merged.text = 'vertical span';

      expect(merged.gridSpan, 1);
      expect(table.cell(1, 0).text, 'vertical span');
      expect(table.cell(2, 0).text, 'vertical span');
    });
  });
}

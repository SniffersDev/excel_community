import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

import 'streaming_fixture.dart';

// The import path of the SFEMP platform reads whole workbooks through
// ExcelStreamReader. These tests pin what that needs beyond the first
// version of the reader: cell types the writer never produces (inline and
// formula strings), dates by cell style, sparse rows, an async variant that
// yields between slices, and the row count from <dimension>.

const _styles = '<numFmts count="2">'
    '<numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>'
    '<numFmt numFmtId="165" formatCode="0.00"/>'
    '</numFmts>'
    '<cellXfs count="5">'
    '<xf numFmtId="0"/>'
    '<xf numFmtId="14"/>'
    '<xf numFmtId="164"/>'
    '<xf numFmtId="165"/>'
    '<xf numFmtId="22"/>'
    '</cellXfs>';

void main() {
  group('ExcelStreamReader cell types', () {
    test('reads shared, inline and formula strings, numbers and bools', () {
      final bytes = buildXlsx(
        sharedStringsXml:
            '<si><t>Unit A</t></si><si><r><t>Rich </t></r><r><t>text</t></r></si>',
        sheetXml: {
          'Units': '<row r="1">'
              '<c r="A1" t="s"><v>0</v></c>'
              '<c r="B1" t="s"><v>1</v></c>'
              '<c r="C1" t="inlineStr"><is><t>inline &amp; more</t></is></c>'
              '<c r="D1"><v>42</v></c>'
              '<c r="E1"><v>4.5</v></c>'
              '<c r="F1" t="b"><v>1</v></c>'
              '<c r="G1" t="str"><f>A1&amp;B1</f><v>calc</v></c>'
              '<c r="H1" t="n"><v>7</v></c>'
              '</row>',
        },
      );

      final rows = ExcelStreamReader.fromBytes(bytes).readSheet('Units').toList();

      expect(rows, hasLength(1));
      expect(rows.single.cells.map((c) => c.value), [
        TextCellValue('Unit A'),
        TextCellValue('Rich text'),
        TextCellValue('inline & more'),
        IntCellValue(42),
        DoubleCellValue(4.5),
        BoolCellValue(true),
        TextCellValue('calc'),
        IntCellValue(7),
      ]);
    });

    test('places cells by their reference', () {
      final bytes = buildXlsx(
        sheetXml: {
          'S': '<row r="2"><c r="D2"><v>4</v></c><c r="F2"><v>6</v></c></row>',
        },
      );

      final row = ExcelStreamReader.fromBytes(bytes).readSheet('S').single;

      expect(row.index, 1);
      expect(row.cells.map((c) => c.columnIndex), [3, 5]);
      expect(row.cellAt(0), isNull);
      expect(row.cellAt(3), IntCellValue(4));
    });

    test('reads dates by style', () {
      final bytes = buildXlsx(
        stylesXml: _styles,
        sheetXml: {
          'S': '<row r="1">'
              '<c r="A1" s="1"><v>45658</v></c>'
              '<c r="B1" s="2"><v>45658</v></c>'
              '<c r="C1" s="3"><v>45658</v></c>'
              '<c r="D1" s="4"><v>45658.5</v></c>'
              '<c r="E1" s="0"><v>45658</v></c>'
              '<c r="F1"><v>45658</v></c>'
              '</row>',
        },
      );

      final row = ExcelStreamReader.fromBytes(bytes).readSheet('S').single;

      expect(row.cellAt(0), const DateCellValue(year: 2025, month: 1, day: 1));
      expect(row.cellAt(1), const DateCellValue(year: 2025, month: 1, day: 1));
      expect(row.cellAt(2), IntCellValue(45658));
      expect(
        row.cellAt(3),
        const DateTimeCellValue(
            year: 2025, month: 1, day: 1, hour: 12, minute: 0, second: 0),
      );
      expect(row.cellAt(4), IntCellValue(45658));
      expect(row.cellAt(5), IntCellValue(45658));
    });

    test('reads the dates the DOM writer saves', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'),
          const DateCellValue(year: 2025, month: 1, day: 1));
      sheet.updateCell(
          CellIndex.indexByString('B1'),
          const DateTimeCellValue(
              year: 2025, month: 1, day: 1, hour: 12, minute: 30, second: 0));
      sheet.updateCell(CellIndex.indexByString('C1'), IntCellValue(45658));
      final bytes = excel.encode()!;

      final row = ExcelStreamReader.fromBytes(bytes).readSheet('Sheet1').single;

      expect(row.cellAt(0), const DateCellValue(year: 2025, month: 1, day: 1));
      expect(
        row.cellAt(1),
        const DateTimeCellValue(
            year: 2025, month: 1, day: 1, hour: 12, minute: 30, second: 0),
      );
      expect(row.cellAt(2), IntCellValue(45658));
    });

    test('self-closing and empty rows are yielded without cells', () {
      final bytes = buildXlsx(
        sheetXml: {
          'S': '<row r="1"/><row r="2"></row><row r="3"><c r="A3"><v>1</v></c></row>',
        },
      );

      final rows = ExcelStreamReader.fromBytes(bytes).readSheet('S').toList();

      expect(rows.map((r) => r.index), [0, 1, 2]);
      expect(rows[0].cells, isEmpty);
      expect(rows[1].cells, isEmpty);
      expect(rows[2].cells, hasLength(1));
    });
  });

  group('ExcelStreamReader.readSheetAsync', () {
    String bigSheet(int rows) {
      final buf = StringBuffer();
      for (var r = 1; r <= rows; r++) {
        buf.write('<row r="$r"><c r="A$r"><v>$r</v></c>'
            '<c r="B$r" t="inlineStr"><is><t>row $r text, padded to make the row longer</t></is></c></row>');
      }
      return buf.toString();
    }

    test('yields the same rows as readSheet and awaits the pause', () async {
      final bytes = buildXlsx(sheetXml: {'S': bigSheet(5000)});
      final reader = ExcelStreamReader.fromBytes(bytes);
      var pauses = 0;

      final sync = reader.readSheet('S').toList();
      final async = await reader
          .readSheetAsync('S', pause: () async => pauses++)
          .toList();

      expect(async.length, sync.length);
      expect(async.length, 5000);
      for (var i = 0; i < sync.length; i++) {
        expect(async[i].index, sync[i].index);
        expect(async[i].cells.map((c) => c.value),
            sync[i].cells.map((c) => c.value));
      }
      expect(pauses, greaterThan(0));
    });

    test('a tiny slice size keeps every row intact', () async {
      final bytes = buildXlsx(sheetXml: {'S': bigSheet(300)});

      final rows = await ExcelStreamReader.fromBytes(bytes)
          .readSheetAsync('S', sliceSize: 100)
          .toList();

      expect(rows, hasLength(300));
      expect(rows.last.cellAt(0), IntCellValue(300));
      expect(rows.last.cellAt(1),
          TextCellValue('row 300 text, padded to make the row longer'));
    });

    test('works without a pause callback', () async {
      final bytes = buildXlsx(sheetXml: {'S': bigSheet(10)});

      final rows =
          await ExcelStreamReader.fromBytes(bytes).readSheetAsync('S').toList();

      expect(rows, hasLength(10));
    });
  });

  group('ExcelStreamReader.rowCount', () {
    test('reads the dimension', () {
      final bytes = buildXlsx(
        sheetXml: {'A': '<row r="1"/>', 'B': '<row r="1"/>'},
        dimension: {'A': 'A1:B1001'},
      );
      final reader = ExcelStreamReader.fromBytes(bytes);

      expect(reader.rowCount('A'), 1001);
      expect(reader.rowCount('B'), isNull);
    });

    test('unknown sheet throws like readSheet', () {
      final bytes = buildXlsx(sheetXml: {'A': '<row r="1"/>'});

      expect(() => ExcelStreamReader.fromBytes(bytes).rowCount('Nope'),
          throwsArgumentError);
    });
  });
}

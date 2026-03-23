import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('ExcelStreamWriter', () {
    test('basic write and read back via DOM Excel', () {
      final writer = ExcelStreamWriter(sheetName: 'Data');
      writer.addHeaderRow(['Name', 'Age', 'Score']);
      writer.addRow([
        TextCellValue('Alice'),
        IntCellValue(30),
        DoubleCellValue(95.5),
      ]);
      writer.addRow([
        TextCellValue('Bob'),
        IntCellValue(25),
        DoubleCellValue(87.3),
      ]);

      final bytes = writer.encode();
      expect(bytes.length, greaterThan(0));

      // Read back with DOM-based Excel
      final excel = Excel.decodeBytes(bytes);
      final sheet = excel.tables['Data']!;
      expect(sheet.maxRows, equals(3)); // header + 2 data rows
      expect(sheet.maxColumns, equals(3));

      // Verify header
      expect(sheet.rows[0][0]!.value.toString(), equals('Name'));
      expect(sheet.rows[0][1]!.value.toString(), equals('Age'));
      expect(sheet.rows[0][2]!.value.toString(), equals('Score'));

      // Verify data
      expect(sheet.rows[1][0]!.value.toString(), equals('Alice'));
      expect(sheet.rows[2][0]!.value.toString(), equals('Bob'));
    });

    test('multi-sheet support', () {
      final writer = ExcelStreamWriter(sheetName: 'Sales');
      writer.addHeaderRow(['Month', 'Revenue']);
      writer.addRow([TextCellValue('Jan'), IntCellValue(1000)]);
      writer.addRow([TextCellValue('Feb'), IntCellValue(1200)]);

      writer.addSheet('Inventory');
      writer.addHeaderRow(['SKU', 'Qty']);
      writer.addRow([TextCellValue('A001'), IntCellValue(500)]);

      final bytes = writer.encode();

      final excel = Excel.decodeBytes(bytes);
      expect(excel.tables.keys, containsAll(['Sales', 'Inventory']));

      final sales = excel.tables['Sales']!;
      expect(sales.maxRows, equals(3));
      expect(sales.rows[1][0]!.value.toString(), equals('Jan'));

      final inventory = excel.tables['Inventory']!;
      expect(inventory.maxRows, equals(2));
      expect(inventory.rows[1][0]!.value.toString(), equals('A001'));
    });

    test('shared strings are shared across sheets', () {
      final writer = ExcelStreamWriter(sheetName: 'Sheet1');
      writer.addRow([TextCellValue('Hello')]);

      writer.addSheet('Sheet2');
      writer.addRow([TextCellValue('Hello')]); // same string

      final bytes = writer.encode();
      final excel = Excel.decodeBytes(bytes);

      expect(
          excel.tables['Sheet1']!.rows[0][0]!.value.toString(), equals('Hello'));
      expect(
          excel.tables['Sheet2']!.rows[0][0]!.value.toString(), equals('Hello'));
    });

    test('encode can only be called once', () {
      final writer = ExcelStreamWriter();
      writer.addRow([IntCellValue(1)]);
      writer.encode();
      expect(() => writer.encode(), throwsStateError);
      expect(() => writer.addRow([IntCellValue(2)]), throwsStateError);
    });

    test('10k rows performance', () {
      final writer = ExcelStreamWriter();
      const rows = 10000;
      const cols = 20;

      final sw = Stopwatch()..start();

      writer
          .addHeaderRow(List.generate(cols, (c) => 'Col_$c'));

      for (var r = 0; r < rows; r++) {
        writer.addRow(List.generate(cols, (c) {
          switch (c % 4) {
            case 0:
              return TextCellValue('R${r}C$c');
            case 1:
              return IntCellValue(r * cols + c);
            case 2:
              return DoubleCellValue(r * 1.5 + c * 0.1);
            default:
              return IntCellValue(r + c);
          }
        }));
      }

      final genTime = sw.elapsed;
      sw.reset();

      final bytes = writer.encode();
      final encodeTime = sw.elapsed;

      print('StreamWriter 10k×20: gen=${genTime.inMilliseconds}ms, '
          'encode=${encodeTime.inMilliseconds}ms, '
          'size=${(bytes.length / 1024).toStringAsFixed(0)}KB');

      expect(bytes.length, greaterThan(0));
    });
  });

  group('ExcelStreamReader', () {
    test('read back stream-written file', () {
      final writer = ExcelStreamWriter(sheetName: 'TestSheet');
      writer.addHeaderRow(['A', 'B', 'C']);
      writer.addRow([
        TextCellValue('hello'),
        IntCellValue(42),
        DoubleCellValue(3.14),
      ]);
      writer.addRow([
        TextCellValue('world'),
        IntCellValue(99),
        DoubleCellValue(2.71),
      ]);

      final bytes = writer.encode();

      final reader = ExcelStreamReader.fromBytes(bytes);
      expect(reader.sheetNames, equals(['TestSheet']));

      final rows = reader.readSheet('TestSheet').toList();
      expect(rows.length, equals(3));

      // Header row
      expect(rows[0].index, equals(0));
      expect(rows[0].cells[0].value, equals(TextCellValue('A')));
      expect(rows[0].cells[1].value, equals(TextCellValue('B')));
      expect(rows[0].cells[2].value, equals(TextCellValue('C')));

      // Data row 1
      expect(rows[1].cells[0].value, equals(TextCellValue('hello')));
      expect(rows[1].cells[1].value, equals(IntCellValue(42)));

      // Data row 2
      expect(rows[2].cells[0].value, equals(TextCellValue('world')));
    });

    test('read DOM-created file via stream reader', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(
          CellIndex.indexByString('A1'), TextCellValue('Header'));
      sheet.updateCell(
          CellIndex.indexByString('B1'), TextCellValue('Value'));
      sheet.updateCell(
          CellIndex.indexByString('A2'), TextCellValue('Row1'));
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(100));

      final bytes = excel.encode()!;

      final reader = ExcelStreamReader.fromBytes(bytes);
      expect(reader.sheetNames, contains('Sheet1'));

      final rows = reader.readSheet('Sheet1').toList();
      expect(rows.length, equals(2));
      expect(rows[0].cells[0].value, equals(TextCellValue('Header')));
      expect(rows[1].cells[1].value, equals(IntCellValue(100)));
    });

    test('multi-sheet reading', () {
      final writer = ExcelStreamWriter(sheetName: 'Alpha');
      writer.addRow([TextCellValue('a1')]);

      writer.addSheet('Beta');
      writer.addRow([TextCellValue('b1')]);

      final bytes = writer.encode();

      final reader = ExcelStreamReader.fromBytes(bytes);
      expect(reader.sheetNames, containsAll(['Alpha', 'Beta']));

      final alphaRows = reader.readSheet('Alpha').toList();
      expect(alphaRows[0].cells[0].value, equals(TextCellValue('a1')));

      final betaRows = reader.readSheet('Beta').toList();
      expect(betaRows[0].cells[0].value, equals(TextCellValue('b1')));
    });

    test('invalid sheet name throws', () {
      final writer = ExcelStreamWriter();
      writer.addRow([IntCellValue(1)]);
      final bytes = writer.encode();

      final reader = ExcelStreamReader.fromBytes(bytes);
      expect(
          () => reader.readSheet('Nonexistent').toList(), throwsArgumentError);
    });

    test('cellAt convenience method', () {
      final row = ExcelRow(0, [
        ExcelCell(0, TextCellValue('A')),
        ExcelCell(2, TextCellValue('C')),
      ]);
      expect(row.cellAt(0), equals(TextCellValue('A')));
      expect(row.cellAt(1), isNull); // no cell at col 1
      expect(row.cellAt(2), equals(TextCellValue('C')));
    });
  });

  group('Stream round-trip 400k', () {
    test('400k rows write via stream', () {
      const rows = 400000;
      const cols = 10;

      final sw = Stopwatch()..start();
      final writer = ExcelStreamWriter();
      writer
          .addHeaderRow(List.generate(cols, (c) => 'H$c'));

      for (var r = 0; r < rows; r++) {
        writer.addRow(List.generate(cols, (c) {
          switch (c % 3) {
            case 0:
              return TextCellValue('R${r}C$c');
            case 1:
              return IntCellValue(r * cols + c);
            default:
              return DoubleCellValue(r * 1.5);
          }
        }));
      }

      final genTime = sw.elapsed;
      sw.reset();

      final bytes = writer.encode();
      final encodeTime = sw.elapsed;

      print('StreamWriter 400k×$cols: gen=${genTime.inSeconds}s, '
          'encode=${encodeTime.inSeconds}s, '
          'size=${(bytes.length / 1024 / 1024).toStringAsFixed(1)}MB');

      expect(bytes.length, greaterThan(0));
      expect(genTime.inMinutes, lessThan(2),
          reason: '400k stream write should be fast');
    }, timeout: Timeout(Duration(minutes: 5)));
  });
}

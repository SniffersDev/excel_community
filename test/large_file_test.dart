import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  test('400,000 rows x 20 columns - large file stress test', () {
    const rows = 400000;
    const cols = 20;

    final stopwatch = Stopwatch()..start();

    final excel = Excel.createExcel();
    final sheet = excel['Sheet1'];

    // Header row
    for (var c = 0; c < cols; c++) {
      sheet
          .cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: c))
          .value = TextCellValue('Header_$c');
    }

    // Data rows with mixed types
    for (var r = 1; r <= rows; r++) {
      for (var c = 0; c < cols; c++) {
        final cell = sheet
            .cell(CellIndex.indexByColumnRow(rowIndex: r, columnIndex: c));
        switch (c % 4) {
          case 0:
            cell.value = TextCellValue('Row${r}_Col$c');
            break;
          case 1:
            cell.value = IntCellValue(r * cols + c);
            break;
          case 2:
            cell.value = DoubleCellValue(r * 1.5 + c * 0.1);
            break;
          case 3:
            cell.value = IntCellValue(r + c);
            break;
        }
      }
    }

    final generateTime = stopwatch.elapsed;
    print('400k×20: Data generation took $generateTime');
    stopwatch.reset();

    // Encode
    final bytes = excel.encode();
    final encodeTime = stopwatch.elapsed;
    print(
        '400k×20: Encode took $encodeTime (${bytes!.length} bytes, ${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB)');
    stopwatch.reset();

    expect(bytes, isNotNull);
    expect(bytes.length, greaterThan(0));

    // Round-trip decode
    final excel2 = Excel.decodeBytes(bytes);
    final decodeTime = stopwatch.elapsed;
    print('400k×20: Decode took $decodeTime');

    final sheet2 = excel2.tables['Sheet1']!;
    expect(sheet2.maxRows, equals(rows + 1)); // +1 for header
    expect(sheet2.maxColumns, equals(cols));

    // Spot-check header
    expect(sheet2.rows[0][0]!.value.toString(), equals('Header_0'));
    expect(sheet2.rows[0][19]!.value.toString(), equals('Header_19'));

    // Spot-check data
    expect(sheet2.rows[1][0]!.value, equals(TextCellValue('Row1_Col0')));
    expect(sheet2.rows[1][1]!.value, equals(IntCellValue(1 * cols + 1)));
    expect(
        sheet2.rows[400000][0]!.value, equals(TextCellValue('Row400000_Col0')));

    expect(encodeTime.inMinutes, lessThan(5),
        reason: '400k encode should complete under 5 minutes');
  });
}

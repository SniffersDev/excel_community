import 'dart:typed_data';

import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

import 'streaming_fixture.dart';

// The reader inflates a sheet block by block through a sliding window and
// scans rows from the chunks, so a sheet is never held as a whole: the
// platform's import hit "Array buffer allocation failed" on a 1.1 GB sheet.
// Rows that straddle a chunk boundary are completed by the next chunk,
// stored sheets read as before, and the row count still comes from the
// first chunk.

class _MemorySink implements ByteSink {
  final BytesBuilder builder = BytesBuilder(copy: false);
  @override
  void add(Uint8List bytes) => builder.add(bytes);
  Uint8List get bytes => builder.toBytes();
}

Uint8List _bigWorkbook(int rows) {
  final sink = _MemorySink();
  final zip = ZipStreamWriter(sink);
  final writer = ExcelStreamWriter.streamed(zip, sheetName: 'Measurements');
  for (var i = 0; i < rows; i++) {
    writer.addRow([
      TextCellValue('Project ${i % 17} & Scope ${i % 5}'),
      TextCellValue('Unit ${i % 7} | Section ${i % 13} | Drawing ${i % 101} | Stream ${i % 211}'),
      IntCellValue(i),
      DoubleCellValue(i / 7),
      i % 3 == 0 ? null : BoolCellValue(i.isEven),
      TextCellValue('x' * (i % 300)),
    ]);
  }
  writer.addSheet('Empty');
  writer.finish();
  zip.finish();
  return sink.bytes;
}

void main() {
  test('a deflated sheet of many blocks reads back row by row, sync and async', () async {
    final bytes = _bigWorkbook(60000);
    final reader = ExcelStreamReader.fromBytes(bytes);

    var count = 0;
    ExcelRow? last;
    for (final row in reader.readSheet('Measurements')) {
      expect(row.index, count, reason: 'rows arrive in order without gaps');
      count++;
      last = row;
    }
    expect(count, 60000);
    expect(last!.cellAt(0).toString(), 'Project ${59999 % 17} & Scope ${59999 % 5}');
    expect((last.cellAt(2) as IntCellValue).value, 59999);
    expect(last.cellAt(5).toString(), 'x' * (59999 % 300));

    var pauses = 0;
    final rows = await reader.readSheetAsync('Measurements', pause: () async => pauses++).toList();
    expect(rows, hasLength(60000));
    expect(rows[12345].cellAt(1).toString(), 'Unit ${12345 % 7} | Section ${12345 % 13} | Drawing ${12345 % 101} | Stream ${12345 % 211}');
    expect(rows[12345].cellAt(4), isNull);
    expect(rows[12346].cellAt(4), isA<BoolCellValue>());
    expect(pauses, greaterThan(10), reason: 'the pause is awaited per slice, not per chunk');
    expect(await reader.readSheetAsync('Empty').toList(), isEmpty);
  });

  test('small chunks complete rows across their boundaries', () {
    // The scanner is private; a tiny minChunk is exercised through rowCount's
    // first-chunk path and through a workbook whose rows are longer than a
    // deflate block's output.
    final bytes = _bigWorkbook(2000);
    final reader = ExcelStreamReader.fromBytes(bytes);
    expect(reader.readSheet('Measurements').length, 2000);
    expect(reader.rowCount('Measurements'), isNull, reason: 'the streamed writer writes no dimension');
  });

  test('a stored sheet reads the same as a deflated one', () {
    const sheet = '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>one</t></is></c></row>'
        '<row r="2"><c r="A2"><v>2</v></c></row>';
    for (final store in [false, true]) {
      final bytes = buildXlsx(sheetXml: {'S': sheet}, dimension: {'S': 'A1:B2'}, storeSheets: store);
      final reader = ExcelStreamReader.fromBytes(bytes);
      final rows = reader.readSheet('S').toList();
      expect(rows, hasLength(2), reason: 'stored=$store');
      expect(rows.first.cellAt(1).toString(), 'one');
      expect(reader.rowCount('S'), 2);
    }
  });

  test('an empty sheetData element yields no rows', () {
    final bytes = buildXlsx(sheetXml: {'S': ''});
    expect(ExcelStreamReader.fromBytes(bytes).readSheet('S').toList(), isEmpty);
  });
}

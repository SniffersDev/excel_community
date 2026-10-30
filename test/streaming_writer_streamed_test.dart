import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

// ExcelStreamWriter.streamed writes each sheet into a ZipStreamWriter while
// the rows arrive, text cells inline, the workbook parts last; a workbook of
// any length costs the memory of the pending rows only, and the result reads
// back through ExcelStreamReader and opens as a normal xlsx.

class _MemorySink implements ByteSink {
  final BytesBuilder builder = BytesBuilder(copy: false);
  int adds = 0;
  @override
  void add(Uint8List bytes) {
    adds++;
    builder.add(bytes);
  }

  int get length => builder.length;
  Uint8List get bytes => builder.toBytes();
}

List<CellValue?> _row(int i) => [
      TextCellValue('name $i & <tag> "q" \'a\''),
      IntCellValue(i),
      DoubleCellValue(i / 8),
      BoolCellValue(i.isEven),
      i % 3 == 0 ? null : TextCellValue('x'),
    ];

/// The names of the entries whose local header carries a zip64 extra field.
List<String> _zip64LocalEntries(Uint8List bytes) {
  final data = ByteData.sublistView(bytes);
  final names = <String>[];
  for (final file in ZipDecoder().decodeBytes(bytes).files) {
    final at = (file.rawContent as ZipFile).header!.localHeaderOffset;
    final nameLength = data.getUint16(at + 26, Endian.little);
    final extraLength = data.getUint16(at + 28, Endian.little);
    var extraAt = at + 30 + nameLength;
    final end = extraAt + extraLength;
    while (extraAt + 4 <= end) {
      if (data.getUint16(extraAt, Endian.little) == 0x0001) names.add(file.name);
      extraAt += 4 + data.getUint16(extraAt + 2, Endian.little);
    }
  }
  return names;
}

void main() {
  // Excel asks to repair a workbook whose parts announce zip64 when they do
  // not need it; a sheet Excel can open never passes 4 GiB.
  test('no part of a streamed workbook is written as zip64', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    final writer = ExcelStreamWriter.streamed(zip, sheetName: 'Units', flushRows: 100);
    writer.addHeaderRow(['name', 'n', 'd', 'b', 'x']);
    for (var i = 0; i < 500; i++) {
      writer.addRow(_row(i));
    }
    writer.addSheet('Sources');
    writer.addRow(_row(1));
    writer.finish();
    zip.finish();

    expect(_zip64LocalEntries(sink.bytes), isEmpty);
  });

  test('a streamed workbook round-trips through ExcelStreamReader', () async {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    final writer = ExcelStreamWriter.streamed(zip, sheetName: 'Units', flushRows: 100);
    writer.addHeaderRow(['name', 'n', 'd', 'b', 'x']);
    for (var i = 0; i < 1000; i++) {
      writer.addRow(_row(i));
    }
    writer.addSheet('Sources');
    writer.addHeaderRow(['name', 'n', 'd', 'b', 'x'], headerStyle: CellStyle(bold: true));
    for (var i = 0; i < 10000; i++) {
      writer.addRow(_row(i));
    }
    writer.addSheet('Empty');
    writer.finish();
    zip.finish();

    final reader = ExcelStreamReader.fromBytes(sink.bytes);
    expect(reader.sheetNames, ['Units', 'Sources', 'Empty']);
    final rows = await reader.readSheetAsync('Sources').toList();
    expect(rows, hasLength(10001));
    expect(rows.first.cellAt(0).toString(), 'name');
    // Data row 6999: a multiple of 3, so its last cell is empty.
    final row = rows[7000];
    expect(row.index, 7000);
    expect(row.cellAt(0).toString(), 'name 6999 & <tag> "q" \'a\'');
    expect((row.cellAt(1) as IntCellValue).value, 6999);
    expect((row.cellAt(2) as DoubleCellValue).value, 6999 / 8);
    expect((row.cellAt(3) as BoolCellValue).value, isFalse);
    expect(row.cellAt(4), isNull);
    expect(rows[7001].cellAt(4).toString(), 'x');
    expect(await reader.readSheetAsync('Empty').toList(), isEmpty);
    expect((await reader.readSheetAsync('Units').toList()), hasLength(1001));
  });

  test('sheet bytes reach the zip before the sheet ends, and the parts come last', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    final writer = ExcelStreamWriter.streamed(zip, flushRows: 50);
    final headerOnly = sink.length;
    // Flushed rows go through the deflater, which emits its first block once
    // it has gathered enough input; well before the sheet ends.
    for (var i = 0; i < 5000; i++) {
      writer.addRow(_row(i));
    }
    expect(sink.length, greaterThan(headerOnly + 1000), reason: 'sheet data reached the sink before finish');
    writer.finish();
    zip.finish();

    final archive = ZipDecoder().decodeBytes(sink.bytes, verify: true);
    expect(archive.map((f) => f.name).toList(), [
      'xl/worksheets/sheet1.xml',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/sharedStrings.xml',
      '_rels/.rels',
      '[Content_Types].xml',
    ]);
  });

  test('inline strings leave the shared strings part empty', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    final writer = ExcelStreamWriter.streamed(zip);
    writer.addHeaderRow(['a']);
    writer.addRow([TextCellValue('same'), TextCellValue('same')]);
    writer.finish();
    zip.finish();

    final archive = ZipDecoder().decodeBytes(sink.bytes, verify: true);
    final shared = utf8.decode(archive.findFile('xl/sharedStrings.xml')!.content);
    expect(shared, contains('count="0"'));
    final sheet = utf8.decode(archive.findFile('xl/worksheets/sheet1.xml')!.content);
    expect(sheet, contains('t="inlineStr"'));
    expect(sheet, contains('<is><t xml:space="preserve">same</t></is>'));
  });

  test('shared strings can still be asked for in streamed mode', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    final writer = ExcelStreamWriter.streamed(zip, inlineStrings: false);
    writer.addRow([TextCellValue('same'), TextCellValue('same')]);
    writer.finish();
    zip.finish();

    final archive = ZipDecoder().decodeBytes(sink.bytes, verify: true);
    expect(utf8.decode(archive.findFile('xl/sharedStrings.xml')!.content), contains('uniqueCount="1"'));
    expect(utf8.decode(archive.findFile('xl/worksheets/sheet1.xml')!.content), contains('t="s"'));
  });

  test('encode() and finish() belong to their own modes', () {
    final streamed = ExcelStreamWriter.streamed(ZipStreamWriter(_MemorySink()));
    expect(streamed.encode, throwsStateError);
    streamed.finish();
    expect(streamed.finish, throwsStateError);
    expect(() => streamed.addRow([IntCellValue(1)]), throwsStateError);

    final buffered = ExcelStreamWriter();
    expect(buffered.finish, throwsStateError);
    buffered.addRow([IntCellValue(1)]);
    expect(buffered.encode(), isNotEmpty);
  });
}

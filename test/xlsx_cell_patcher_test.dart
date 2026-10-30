import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

// XlsxCellPatcher copies a workbook (or a zip of workbooks) entry by entry,
// rewriting only the patched cells of the patched sheets as inline strings
// that keep their style; every other entry is the original's, byte for byte.

class _MemorySink implements ByteSink {
  final BytesBuilder builder = BytesBuilder(copy: false);
  @override
  void add(Uint8List bytes) => builder.add(bytes);
  Uint8List get bytes => builder.toBytes();
}

/// A workbook from the DOM writer: shared strings, a styled cell, and two
/// sheets (Units and Notes).
List<int> _workbook() {
  final excel = Excel.createExcel();
  excel.rename('Sheet1', 'Units');
  final units = excel['Units'];
  units.appendRow([TextCellValue('unitName'), TextCellValue('country'), TextCellValue('hours')]);
  units.appendRow([TextCellValue('U1'), TextCellValue('BE'), IntCellValue(10)]);
  units.appendRow([TextCellValue('U2'), TextCellValue('XX'), IntCellValue(20)]);
  units.appendRow([TextCellValue('U3'), null, IntCellValue(30)]);
  units.cell(CellIndex.indexByString('B3')).cellStyle = CellStyle(bold: true);
  final notes = excel['Notes'];
  notes.appendRow([TextCellValue('Read me'), TextCellValue('do not touch')]);
  return excel.save()!;
}

Archive _unzip(List<int> bytes) => ZipDecoder().decodeBytes(bytes);

/// Every cell value of [sheet] by its reference, as the stream reader reads it.
Map<String, String> _cells(List<int> bytes, String sheet) => {
      for (final row in ExcelStreamReader.fromBytes(bytes).readSheet(sheet))
        for (final cell in row.cells) getCellId(cell.columnIndex, row.index): cell.value.toString(),
    };

String _sheetXml(List<int> bytes, String part) => utf8.decode(_unzip(bytes).findFile(part)!.content);

void main() {
  test('only the patched cells change; every other entry is copied byte for byte', () async {
    final original = _workbook();
    final sink = _MemorySink();

    final missed = await XlsxCellPatcher.writeWorkbook(original, sink, const [
      XlsxCellPatch(sheet: 'Units', row: 2, column: 1, text: 'NL & <B>'),
    ]);

    expect(missed, isEmpty);
    final patched = sink.bytes;
    final before = _cells(original, 'Units');
    final after = _cells(patched, 'Units');
    expect(after['B3'], 'NL & <B>');
    expect({...after}..remove('B3'), {...before}..remove('B3'));
    expect(_cells(patched, 'Notes'), _cells(original, 'Notes'));

    // The patched cell is an inline string with the original style.
    const part = 'xl/worksheets/sheet1.xml';
    final style = RegExp(r'<c r="B3"[^>]*\ss="(\d+)"').firstMatch(_sheetXml(original, part))!.group(1);
    expect(
      _sheetXml(patched, part),
      contains('<c r="B3" s="$style" t="inlineStr"><is><t>NL &amp; &lt;B&gt;</t></is></c>'),
    );

    // Every other entry: same name, order, CRC and raw (compressed) bytes.
    final a = _unzip(original).files;
    final b = _unzip(patched).files;
    expect(b.map((f) => f.name), a.map((f) => f.name));
    for (var i = 0; i < a.length; i++) {
      if (a[i].name == part) continue;
      final ra = a[i].rawContent as ZipFile;
      final rb = b[i].rawContent as ZipFile;
      expect(rb.crc32, ra.crc32, reason: a[i].name);
      expect(rb.compressionMethod, ra.compressionMethod, reason: a[i].name);
      expect(rb.getRawContent(), ra.getRawContent(), reason: a[i].name);
    }

    // And it still opens as a workbook.
    final decoded = Excel.decodeBytes(patched);
    expect(decoded['Units'].cell(CellIndex.indexByString('B3')).value.toString(), 'NL & <B>');
  });

  // Excel asks to repair a workbook whose parts announce zip64 when they do
  // not need it.
  test('the patched sheet is not written as zip64', () async {
    final sink = _MemorySink();

    await XlsxCellPatcher.writeWorkbook(_workbook(), sink, const [
      XlsxCellPatch(sheet: 'Units', row: 2, column: 1, text: 'NL'),
    ]);

    final bytes = sink.bytes;
    final data = ByteData.sublistView(bytes);
    for (final file in _unzip(bytes).files) {
      final at = (file.rawContent as ZipFile).header!.localHeaderOffset;
      final nameLength = data.getUint16(at + 26, Endian.little);
      final extraLength = data.getUint16(at + 28, Endian.little);
      var extraAt = at + 30 + nameLength;
      while (extraAt + 4 <= at + 30 + nameLength + extraLength) {
        expect(data.getUint16(extraAt, Endian.little), isNot(0x0001), reason: file.name);
        extraAt += 4 + data.getUint16(extraAt + 2, Endian.little);
      }
    }
  });

  test('a missing cell is inserted in column order; an empty patch empties the cell', () async {
    final sink = _MemorySink();

    await XlsxCellPatcher.writeWorkbook(_workbook(), sink, const [
      XlsxCellPatch(sheet: 'Units', row: 3, column: 1, text: 'FR'),
      XlsxCellPatch(sheet: 'Units', row: 1, column: 1, text: ''),
    ]);

    final cells = _cells(sink.bytes, 'Units');
    expect([cells['A4'], cells['B4'], cells['C4']], ['U3', 'FR', '30']);
    expect(cells.containsKey('B2'), isFalse);
    expect(cells['A2'], 'U1');
    final xml = _sheetXml(sink.bytes, 'xl/worksheets/sheet1.xml');
    expect(xml.indexOf('r="A4"'), lessThan(xml.indexOf('r="B4"')));
    expect(xml.indexOf('r="B4"'), lessThan(xml.indexOf('r="C4"')));
  });

  test('a sheet of many chunks is patched row by row, wherever the rows sit', () async {
    // About 1.5 MB of sheet XML: several inflate chunks.
    final zipSink = _MemorySink();
    final zip = ZipStreamWriter(zipSink);
    final writer = ExcelStreamWriter.streamed(zip, sheetName: 'Measurements', flushRows: 500);
    writer.addHeaderRow(['name', 'value', 'note']);
    for (var i = 0; i < 20000; i++) {
      writer.addRow([TextCellValue('row $i'), IntCellValue(i), TextCellValue('n' * 30)]);
    }
    writer.finish();
    zip.finish();
    final rows = [2, 7777, 13001, 20000];
    final sink = _MemorySink();

    final missed = await XlsxCellPatcher.writeWorkbook(zipSink.bytes, sink, [
      for (final row in rows) XlsxCellPatch(sheet: 'Measurements', row: row, column: 0, text: 'fixed $row'),
      const XlsxCellPatch(sheet: 'Measurements', row: 50000, column: 0, text: 'no such row'),
      const XlsxCellPatch(sheet: 'Nope', row: 1, column: 0, text: 'no such sheet'),
    ]);

    expect(
      missed.map((missed) => (missed.patch.text, missed.reason)),
      unorderedEquals([
        ('no such row', XlsxPatchMissReason.noRow),
        ('no such sheet', XlsxPatchMissReason.noSheet),
      ]),
    );
    var count = 0;
    for (final row in ExcelStreamReader.fromBytes(sink.bytes).readSheet('Measurements')) {
      count++;
      if (row.index == 0) continue; // the header
      final first = row.cells.firstWhere((cell) => cell.columnIndex == 0).value.toString();
      expect(first, rows.contains(row.index) ? 'fixed ${row.index}' : 'row ${row.index - 1}');
    }
    expect(count, 20001);
  });

  test('a package copies every entry and patches only its named workbooks', () async {
    final workbook = _workbook();
    final package = Archive()
      ..addFile(ArchiveFile('site/readme.txt', 5, utf8.encode('hello')))
      ..addFile(ArchiveFile('site/main.xlsx', workbook.length, workbook))
      ..addFile(ArchiveFile('site/other.xlsx', workbook.length, workbook));
    final zipped = _unzip(ZipEncoder().encode(package));
    final sink = _MemorySink();

    final missed = await XlsxCellPatcher.writePackage(zipped, sink, {
      'site/main.xlsx': const [XlsxCellPatch(sheet: 'Units', row: 2, column: 1, text: 'NL')],
      'site/missing.xlsx': const [XlsxCellPatch(sheet: 'Units', row: 2, column: 1, text: 'lost')],
    });

    expect(missed.single.patch.text, 'lost');
    expect(missed.single.reason, XlsxPatchMissReason.noSheet);
    final out = _unzip(sink.bytes);
    expect(out.files.map((f) => f.name), ['site/readme.txt', 'site/main.xlsx', 'site/other.xlsx']);
    expect(utf8.decode(out.findFile('site/readme.txt')!.content), 'hello');
    expect(out.findFile('site/other.xlsx')!.content, workbook);
    final main = out.findFile('site/main.xlsx')!.content;
    expect(_cells(main, 'Units')['B3'], 'NL');
    expect(_cells(main, 'Notes'), _cells(workbook, 'Notes'));
  });

  // Fix round 1: hand-written sheet XML for the edge cases.
  group('sheet XML', () {
    const part = 'xl/worksheets/sheet1.xml';

    /// [_workbook] with the Units sheet's XML replaced by [rows] inside
    /// [sheetData] (or by `<sheetData/>` when [rows] is null).
    List<int> withUnits(String? rows) {
      final xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '${rows == null ? '<sheetData/>' : '<sheetData>$rows</sheetData>'}</worksheet>';
      final out = Archive();
      for (final file in _unzip(_workbook()).files) {
        final content = file.name == part ? utf8.encode(xml) : file.content;
        out.addFile(ArchiveFile(file.name, content.length, content));
      }
      return ZipEncoder().encode(out);
    }

    Future<(List<XlsxMissedPatch>, String)> patch(List<int> workbook, List<XlsxCellPatch> patches) async {
      final sink = _MemorySink();
      final missed = await XlsxCellPatcher.writeWorkbook(workbook, sink, patches);
      return (missed, _sheetXml(sink.bytes, part));
    }

    test('the master of a shared formula is never rewritten; its dependents keep their definition', () async {
      const rows = '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>'
          '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>'
          '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>';
      final workbook = withUnits(rows);

      final (missed, xml) = await patch(workbook, const [
        XlsxCellPatch(sheet: 'Units', row: 0, column: 1, text: 'x'),
        XlsxCellPatch(sheet: 'Units', row: 2, column: 0, text: '30'),
      ]);

      expect(missed.single.patch.row, 0);
      expect(missed.single.reason, XlsxPatchMissReason.formula);
      // The master's row is the original, byte for byte; the dependents keep
      // their shared formula; the other patch is written.
      expect(xml, contains(rows.substring(0, rows.indexOf('<row r="2">'))));
      expect(xml, contains('<c r="B2"><f t="shared" si="0"/><v>4</v></c>'));
      expect(xml, contains('<c r="A3" t="inlineStr"><is><t>30</t></is></c><c r="B3"><f t="shared" si="0"/><v>6</v></c>'));
    });

    test('a plain formula cell is never rewritten', () async {
      const rows = '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="C2"><f>A2+1</f><v>3</v></c></row>';
      final workbook = withUnits(rows);

      final (missed, xml) = await patch(workbook, const [XlsxCellPatch(sheet: 'Units', row: 1, column: 2, text: '9')]);

      expect(missed.single.reason, XlsxPatchMissReason.formula);
      expect(xml, contains(rows));
    });

    test('a row and cells without r are numbered after the ones before them', () async {
      final (missed, xml) = await patch(withUnits('<row r="1"><c r="A1"><v>1</v></c></row><row><c><v>5</v></c><c><v>6</v></c></row>'), const [
        XlsxCellPatch(sheet: 'Units', row: 1, column: 1, text: 'new'),
      ]);

      expect(missed, isEmpty);
      expect(xml, contains('<row><c><v>5</v></c><c r="B2" t="inlineStr"><is><t>new</t></is></c></row>'));
    });

    test('a self-closing row gets its cell and loses its spans', () async {
      final (missed, xml) = await patch(withUnits('<row r="1"><c r="A1"><v>1</v></c></row><row r="2" spans="1:2"/>'), const [
        XlsxCellPatch(sheet: 'Units', row: 1, column: 1, text: 'B'),
      ]);

      expect(missed, isEmpty);
      expect(xml, contains('<row r="2"><c r="B2" t="inlineStr"><is><t>B</t></is></c></row>'));
    });

    test('an empty sheetData has no row to patch', () async {
      final workbook = withUnits(null);

      final (missed, xml) = await patch(workbook, const [XlsxCellPatch(sheet: 'Units', row: 0, column: 0, text: 'x')]);

      expect(missed.single.reason, XlsxPatchMissReason.noRow);
      expect(xml, _sheetXml(workbook, part));
    });

    test('a value with outer spaces keeps them', () async {
      final (missed, xml) = await patch(withUnits('<row r="1"><c r="A1"><v>1</v></c></row>'), const [
        XlsxCellPatch(sheet: 'Units', row: 0, column: 0, text: ' padded '),
      ]);

      expect(missed, isEmpty);
      expect(xml, contains('<c r="A1" t="inlineStr"><is><t xml:space="preserve"> padded </t></is></c>'));
    });
  });
}

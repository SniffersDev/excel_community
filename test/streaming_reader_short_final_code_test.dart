import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

import 'streaming_fixture.dart';

// A sheet whose deflate stream ends right after a short final end-of-block
// code: fewer bits are left than the stream's longest code. The block
// inflater looked ahead for the longest code, met the end of the input and
// threw "Invalid deflate data", although every byte had been decoded; the
// platform's import of an Excel-written workbook failed on this sheet.
// [_sheetDeflated] is that sheet part's raw deflate data, as the workbook
// holds it (a header row of shared strings and empty rows), inline because
// test_resources is gitignored.

const _sheetDeflated =
  'nZfbbuM2EIZfheC9LZGSDzEsL5K4RgO0SLBbdK9paWwTEUlhSJ+26LsvSDlOutuimdzYIvHz18x8EjWcfzqZlh0AvXa24mKYcwa2'
  'do2224rvw2Yw5cwHZRvVOgsVP4Pnnxbzo8NnvwMI7GRa62dY8V0I3SzLfL0Do/zQdWBPpt04NCr4ocNt5jYbXcPS1XsDNmQyz8cZ'
  'QquCdtbvdOf5xc3U77EzCp/33aB2plNBr3WrwzmZvticRKl+djK6RufdJgxrZy4xZb5DUE3KyLTR5Ca7yVR9dcKP2YgyQzjoWNxX'
  'K/lBr9HVS76aFR80G1/NCs5OONvrpuJ/TZfl3WpV3g5WKzkZlHfyfnAjxc1AFstCjEZiIn/J/+bM1LOHrXWo1i1UPJWZnZCdULIT'
  'Fpfg3oPwp7KPM6O05Yt5mnvCxdztQ6stPCHze2MUnu+gdceKC/4y8VlvdyFNZIt5dl3YaAM2JsgQNhW/FbP7cpQ0SfKnhqN/c83i'
  'I7127jkOHpqK52/83opXKf4nZA1s1L4Nn93xV7iEUA5HnKV6zJrzEnwNNlQ8Hxb9jWvX+vTLjI5vG2dGndL/UTdhV3Ep49t2jlUV'
  'nK3Bh5XuU6v3PjjztZf1qV595MVHvvqI4VSMy/gq0O2Ki11xtRPTH+3y/7XL+mRTyZYqqMUc3ZFhD65T8fkQs+K/qrWY11F7G8UV'
  'lwVnoeKeL+aHhcgn8+wQb3AR3f2raPpP0f1FNH4jkvlVk6E7XiOU74kw5njRF0R9SdSPiPoxUT8h6qdE/Q1RL3LqAkFdQEUsqIwF'
  'FbKgUhZUzILKWVBBCyppSSUtqaTTjkpaQCUtqaQllXTctGgLqKQllbSkki6opAsq6YK8bVNJF1TScURbQCVdUEkXVNIFlXRJJV1S'
  'SZdU0iX5C00lHTvL9y3I3nREHWobHrt0/mE7h/qbs0G192ADIPQ96AEw6PqHyR2oeELzabBF3fymLfSjbDHv1BZ+V7jV1rMWNunu'
  'E86wb1LTdXBduhpxtnYhOPMyis6AKV7ONs6F6+Bi/AXCvmOd6gC/6G/Qf91rlXrMPOcsah4xeTTuaP/YgX08AHLmUIMN6bBX8c5h'
  'QKUDZ+tW1c+3tvm606FvLBtUKWbOamjbe2fiadFX3DoLnAGiQ1/xRvuuVWdoOMNZPLzgQ9N3nn0KqxR7HId4THlSGDyr3T7B6Elc'
  'D7GL7w==';

/// The minimal workbook from [buildXlsx] with its sheet replaced by
/// [deflated], written as is: an encoder would deflate the sheet anew.
Uint8List _workbookWithRawSheet(Uint8List deflated, Uint8List inflated) {
  final template = ZipDecoder().decodeBytes(buildXlsx(
    sheetXml: {'CustomLeakThresholds': ''},
    sharedStringsXml: [for (var i = 0; i < 110; i++) '<si><t>s$i</t></si>'].join(),
  ));
  final entries = [
    for (final file in template.files.where((file) => file.isFile))
      file.name == 'xl/worksheets/sheet1.xml'
          ? (name: file.name, method: 8, data: deflated, plain: inflated)
          : (name: file.name, method: 0, data: file.content, plain: file.content),
  ];
  return _zip(entries);
}

Uint8List _zip(List<({String name, int method, Uint8List data, Uint8List plain})> entries) {
  final out = BytesBuilder();
  final central = BytesBuilder();
  void u16(BytesBuilder b, int v) => b.add(Uint8List(2)..buffer.asByteData().setUint16(0, v, Endian.little));
  void u32(BytesBuilder b, int v) => b.add(Uint8List(4)..buffer.asByteData().setUint32(0, v, Endian.little));
  for (final entry in entries) {
    final offset = out.length;
    final name = Uint8List.fromList(entry.name.codeUnits);
    final crc = getCrc32(entry.plain);
    u32(out, 0x04034b50);
    for (final v in [20, 0, entry.method, 0, 0x21]) {
      u16(out, v);
    }
    u32(out, crc);
    u32(out, entry.data.length);
    u32(out, entry.plain.length);
    u16(out, name.length);
    u16(out, 0);
    out
      ..add(name)
      ..add(entry.data);

    u32(central, 0x02014b50);
    for (final v in [20, 20, 0, entry.method, 0, 0x21]) {
      u16(central, v);
    }
    u32(central, crc);
    u32(central, entry.data.length);
    u32(central, entry.plain.length);
    for (final v in [name.length, 0, 0, 0, 0]) {
      u16(central, v);
    }
    u32(central, 0);
    u32(central, offset);
    central.add(name);
  }
  final centralOffset = out.length;
  final centralBytes = central.toBytes();
  out.add(centralBytes);
  u32(out, 0x06054b50);
  for (final v in [0, 0, entries.length, entries.length]) {
    u16(out, v);
  }
  u32(out, centralBytes.length);
  u32(out, centralOffset);
  u16(out, 0);
  return out.toBytes();
}

void main() {
  final deflated = base64.decode(_sheetDeflated);
  // archive's Inflate ignores the failed look-ahead on the final block, so it
  // gives the sheet's content (3,847 bytes, all of it) for the CRC.
  final inflated = Inflate(deflated).getBytes();

  test('the fixture is the whole sheet', () {
    expect(inflated, hasLength(3847));
    expect(String.fromCharCodes(inflated), endsWith('</worksheet>'));
  });

  test('a sheet ending in a code shorter than the longest reads, sync and async', () async {
    final reader = ExcelStreamReader.fromBytes(_workbookWithRawSheet(deflated, inflated));

    final rows = reader.readSheet('CustomLeakThresholds').toList();
    expect(rows, isNotEmpty);
    expect(rows.first.cellAt(0).toString(), 's107');
    expect(rows.first.cellAt(1).toString(), 's108');
    expect(rows.first.cellAt(2).toString(), 's20');

    final asyncRows = await reader.readSheetAsync('CustomLeakThresholds').toList();
    expect(asyncRows, hasLength(rows.length));
  });

  test('a sheet cut short inside its last code still fails', () {
    final truncated = Uint8List.sublistView(deflated, 0, deflated.length - 2);
    final reader = ExcelStreamReader.fromBytes(_workbookWithRawSheet(truncated, inflated));
    expect(() => reader.readSheet('CustomLeakThresholds').toList(), throwsFormatException);
  });
}

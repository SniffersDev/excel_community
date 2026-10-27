import 'dart:convert';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

// ZipStreamWriter writes an archive entry by entry into a sink while the
// content is produced: data descriptors carry the CRC and sizes, deflated
// content can arrive in any number of chunks, and the directory switches to
// zip64 when an entry passes 4 GiB.

class _MemorySink implements ByteSink {
  final BytesBuilder builder = BytesBuilder(copy: false);
  @override
  void add(Uint8List bytes) => builder.add(bytes);
  Uint8List get bytes => builder.toBytes();
}

/// Counts the bytes and keeps only the last [keep] of them.
class _TailSink implements ByteSink {
  _TailSink({this.keep = 4096});
  final int keep;
  int count = 0;
  Uint8List tail = Uint8List(0);
  @override
  void add(Uint8List bytes) {
    count += bytes.length;
    final joined = Uint8List(tail.length + bytes.length)
      ..setRange(0, tail.length, tail)
      ..setRange(tail.length, tail.length + bytes.length, bytes);
    tail = joined.length > keep ? Uint8List.sublistView(joined, joined.length - keep) : joined;
  }
}

Uint8List _text(String s) => Uint8List.fromList(utf8.encode(s));

void main() {
  test('a stored and a deflated entry round-trip through ZipDecoder', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    final big = _text(List.generate(5000, (i) => 'row $i of the sheet').join('\n'));
    zip.addEntry('stored.bin', _text('stored bytes'), method: ZipEntryMethod.stored);
    zip.beginEntry('xl/worksheets/sheet1.xml');
    zip.write(big);
    zip.endEntry();
    zip.addDirectory('docs/empty');
    zip.finish();

    final archive = ZipDecoder().decodeBytes(sink.bytes, verify: true);
    expect(archive.map((f) => f.name), ['stored.bin', 'xl/worksheets/sheet1.xml', 'docs/empty/']);
    expect(utf8.decode(archive.findFile('stored.bin')!.content), 'stored bytes');
    expect(archive.findFile('xl/worksheets/sheet1.xml')!.content, big);
    expect(archive.findFile('xl/worksheets/sheet1.xml')!.compression, CompressionType.deflate);
    expect(archive.findFile('stored.bin')!.compression, CompressionType.none);
    expect(archive.findFile('docs/empty/')!.isFile, isFalse);
    expect(zip.bytesWritten, sink.bytes.length);
    // The deflated entry is smaller than its content.
    expect(sink.bytes.length, lessThan(big.length));
  });

  test('content written in many small chunks equals one write', () {
    final content = _text(List.generate(20000, (i) => 'line $i: ${'x' * (i % 40)}').join('\n'));

    final once = _MemorySink();
    final chunked = _MemorySink();
    ZipStreamWriter(once)
      ..beginEntry('a.txt')
      ..write(content)
      ..endEntry()
      ..finish();
    final zip = ZipStreamWriter(chunked)..beginEntry('a.txt');
    for (var offset = 0; offset < content.length; offset += 777) {
      final end = offset + 777 > content.length ? content.length : offset + 777;
      zip.write(Uint8List.sublistView(content, offset, end));
    }
    zip
      ..endEntry()
      ..finish();

    final decoded = ZipDecoder().decodeBytes(chunked.bytes, verify: true);
    expect(decoded.single.content, content);
    expect(ZipDecoder().decodeBytes(once.bytes, verify: true).single.content, content);
  });

  test('an entry over the 32-bit limit is written as zip64', () {
    final sink = _TailSink();
    final zip = ZipStreamWriter(sink);
    final chunk = Uint8List(16 << 20);
    final total = (1 << 32) + 1;
    zip.beginEntry('huge.bin', method: ZipEntryMethod.stored, large: true);
    var written = 0;
    while (written < total) {
      final size = total - written < chunk.length ? total - written : chunk.length;
      zip.write(size == chunk.length ? chunk : Uint8List.sublistView(chunk, 0, size));
      written += size;
    }
    zip.endEntry();
    zip.addEntry('small.txt', _text('after'), method: ZipEntryMethod.stored);
    zip.finish();

    final tail = sink.tail;
    final data = ByteData.sublistView(tail);
    // Find the central directory record of huge.bin.
    var at = -1;
    for (var i = 0; i + 4 <= tail.length; i++) {
      if (data.getUint32(i, Endian.little) == 0x02014b50) {
        at = i;
        break;
      }
    }
    expect(at, greaterThanOrEqualTo(0), reason: 'central directory not in the tail');
    expect(data.getUint32(at + 20, Endian.little), 0xFFFFFFFF, reason: 'compressed size marker');
    expect(data.getUint32(at + 24, Endian.little), 0xFFFFFFFF, reason: 'uncompressed size marker');
    final nameLength = data.getUint16(at + 28, Endian.little);
    final extraLength = data.getUint16(at + 30, Endian.little);
    expect(utf8.decode(Uint8List.sublistView(tail, at + 46, at + 46 + nameLength)), 'huge.bin');
    final extraAt = at + 46 + nameLength;
    expect(data.getUint16(extraAt, Endian.little), 0x0001);
    expect(extraLength, 20);
    expect(data.getUint64(extraAt + 4, Endian.little), total, reason: 'uncompressed');
    expect(data.getUint64(extraAt + 12, Endian.little), total, reason: 'compressed (stored)');
    // The second entry's offset passed 4 GiB too, so the end record is zip64.
    var eocd = -1;
    for (var i = tail.length - 22; i >= 0; i--) {
      if (data.getUint32(i, Endian.little) == 0x06054b50) {
        eocd = i;
        break;
      }
    }
    expect(eocd, greaterThanOrEqualTo(0));
    expect(data.getUint32(eocd + 16, Endian.little), 0xFFFFFFFF, reason: 'directory offset marker');
    expect(data.getUint32(eocd - 20, Endian.little), 0x07064b50, reason: 'zip64 locator before the end record');
    expect(sink.count, zip.bytesWritten);
    expect(sink.count, greaterThan(total));
  }, timeout: const Timeout(Duration(minutes: 5)));

  test('an entry that passes 4 GiB without large is refused', () {
    final zip = ZipStreamWriter(_TailSink());
    zip.beginEntry('huge.bin', method: ZipEntryMethod.stored);
    final chunk = Uint8List(64 << 20);
    for (var written = 0; written <= (1 << 32); written += chunk.length) {
      zip.write(chunk);
    }
    expect(zip.endEntry, throwsStateError);
  }, timeout: const Timeout(Duration(minutes: 5)));

  test('an abandoned entry is left out of the directory and the rest stays readable', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    zip.addEntry('first.txt', _text('first'));
    zip.beginEntry('broken.bin');
    zip.write(_text('partial content that never completed'));
    zip.abandonEntry();
    zip.addEntry('last.txt', _text('last'), method: ZipEntryMethod.stored);
    zip.finish();

    final archive = ZipDecoder().decodeBytes(sink.bytes, verify: true);
    expect(archive.map((f) => f.name), ['first.txt', 'last.txt']);
    expect(utf8.decode(archive.findFile('last.txt')!.content), 'last');
  });

  test('bytesWritten counts every byte the sink received and misuse throws', () {
    final sink = _MemorySink();
    final zip = ZipStreamWriter(sink);
    expect(() => zip.write(_text('x')), throwsStateError);
    zip.beginEntry('a.txt');
    expect(() => zip.beginEntry('b.txt'), throwsStateError);
    zip.write(_text('hello'));
    zip.endEntry();
    expect(zip.bytesWritten, sink.bytes.length);
    zip.finish();
    expect(zip.bytesWritten, sink.bytes.length);
    expect(zip.finish, throwsStateError);
    expect(() => zip.addEntry('c.txt', _text('late')), throwsStateError);
  });
}

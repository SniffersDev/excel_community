part of excel_community;

/// The inflater's output as a sliding window: every byte written is handed
/// on through [take] and only the last 32 KiB (deflate's longest
/// back-reference) stay behind as history, so a sheet of any size is inflated
/// in bounded memory.
class _WindowOutputStream extends OutputStream {
  _WindowOutputStream() : super(byteOrder: ByteOrder.littleEndian);

  static const int _window = 32 * 1024;
  static const int _initialSize = 256 * 1024;

  Uint8List _buffer = Uint8List(_initialSize);
  int _filled = 0;
  int _taken = 0;
  int _total = 0;

  /// Bytes ever written: the inflater checks back-references against it.
  @override
  int get length => _total;

  void _ensure(int extra) {
    if (_filled + extra <= _buffer.length) return;
    var size = _buffer.length * 2;
    while (size < _filled + extra) {
      size *= 2;
    }
    final grown = Uint8List(size);
    grown.setRange(0, _filled, _buffer);
    _buffer = grown;
  }

  @override
  void writeByte(int value) {
    _ensure(1);
    _buffer[_filled++] = value;
    _total++;
  }

  @override
  void writeBytes(List<int> bytes, {int? length}) {
    final count = length ?? bytes.length;
    _ensure(count);
    _buffer.setRange(_filled, _filled + count, bytes);
    _filled += count;
    _total += count;
  }

  @override
  void writeStream(InputStream stream) => writeBytes(stream.toUint8List());

  /// An LZ77 back-reference from the inflater; copies within the window.
  void writeBackReference(int distance, int count) {
    _ensure(count);
    final src = _filled - distance;
    if (distance >= count) {
      _buffer.setRange(_filled, _filled + count, _buffer, src);
    } else {
      var s = src;
      var d = _filled;
      final end = _filled + count;
      while (d < end) {
        _buffer[d++] = _buffer[s++];
      }
    }
    _filled += count;
    _total += count;
  }

  /// A negative [start] or [end] counts back from the end of what is held.
  @override
  Uint8List subset(int start, [int? end]) {
    final from = start < 0 ? _filled + start : start;
    final to = end == null ? _filled : (end < 0 ? _filled + end : end);
    return Uint8List.sublistView(_buffer, from, to);
  }

  @override
  void clear() {
    _filled = 0;
    _taken = 0;
  }

  @override
  void flush() {}

  /// The bytes written since the last [take], as an owned copy; the window
  /// of history is kept for the back-references that follow.
  Uint8List take() {
    final out = Uint8List.fromList(Uint8List.sublistView(_buffer, _taken, _filled));
    final keep = _filled < _window ? _filled : _window;
    _buffer.setRange(0, keep, _buffer, _filled - keep);
    _filled = keep;
    _taken = keep;
    return out;
  }
}

/// Where each entry's raw stream starts, remembered on the first read.
final Expando<int> _entryStart = Expando<int>();

/// The content of a sheet part in chunks of about [minChunk] bytes: a
/// deflated entry is inflated block by block through [_WindowOutputStream]
/// from its raw bytes, a stored entry comes as it is. The whole sheet is
/// never held, only the zip's bytes and the chunk in hand.
Iterable<Uint8List> _sheetChunks(ArchiveFile file, {int minChunk = 256 * 1024}) sync* {
  final raw = file.rawContent;
  if (raw == null) return;
  // The entry's raw stream is shared and keeps its position, and a reader
  // that stops early (rowCount, a break) never reaches this generator's end:
  // the entry's start is remembered the first time and every read starts
  // from it, so the sheet can be read again and again.
  final input = raw.getStream(decompress: false);
  final savePos = _entryStart[file] ??= input.position;
  input.setPosition(savePos);
  try {
    if (!file.isCompressed) {
      yield input.toUint8List();
      return;
    }
    final output = _WindowOutputStream();
    final inflate = _StreamInflate.lazy(input, output: output);
    final pending = BytesBuilder(copy: false);
    var more = true;
    while (more) {
      more = inflate.inflateBlock();
      pending.add(output.take());
      if (pending.length >= minChunk || !more) {
        if (pending.length > 0) yield pending.takeBytes();
      }
    }
  } finally {
    input.setPosition(savePos);
  }
}

/// Finds `<row>` elements in sheet bytes that arrive in chunks: a chunk may
/// end inside a row, which is kept and completed by the next chunk. The
/// markup is ASCII and a UTF-8 multi-byte sequence never contains an ASCII
/// byte, so searching the bytes is safe; one row at a time is decoded.
class _RowScanner {
  _RowScanner(this._parseCells);

  final List<ExcelCell> Function(String content) _parseCells;

  static final int _keepTail = max(
    ExcelStreamReader._rowOpen.length,
    ExcelStreamReader._sheetDataClose.length,
  );

  Uint8List _pending = Uint8List(0);
  int _base = 0;
  bool _started = false;
  bool _done = false;

  /// No more rows will come (`</sheetData>` or `<sheetData/>` was seen).
  bool get isDone => _done;

  /// The complete rows in [_pending] + [chunk]; [endOffset] is the byte
  /// offset just past the row within the sheet, for pausing.
  List<({ExcelRow row, int endOffset})> feed(Uint8List chunk) {
    if (_done) return const [];
    final bytes = _pending.isEmpty ? chunk : _concat(_pending, chunk);
    final rows = <({ExcelRow row, int endOffset})>[];
    var pos = 0;
    var keepFrom = bytes.length;

    if (!_started) {
      final sheetDataStart = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._sheetDataOpen, 0);
      if (sheetDataStart == -1) {
        _keep(bytes, max(0, bytes.length - ExcelStreamReader._sheetDataOpen.length + 1));
        return rows;
      }
      final afterOpen = ExcelStreamReader._indexOfByte(bytes, ExcelStreamReader._gt, sheetDataStart);
      if (afterOpen == -1) {
        _keep(bytes, sheetDataStart);
        return rows;
      }
      if (bytes[afterOpen - 1] == ExcelStreamReader._slash) {
        _done = true;
        return rows;
      }
      _started = true;
      pos = afterOpen + 1;
    }

    final end = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._sheetDataClose, pos);
    final limit = end == -1 ? bytes.length : end;
    while (pos < limit) {
      final rowStart = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._rowOpen, pos);
      if (rowStart == -1 || rowStart >= limit) {
        keepFrom = end == -1 ? max(pos, bytes.length - _keepTail) : bytes.length;
        break;
      }
      final rowTagEnd = ExcelStreamReader._indexOfByte(bytes, ExcelStreamReader._gt, rowStart);
      if (rowTagEnd == -1) {
        keepFrom = rowStart;
        break;
      }
      final isSelfClosing = bytes[rowTagEnd - 1] == ExcelStreamReader._slash;
      int rowEnd;
      String rowContent;
      if (isSelfClosing) {
        rowEnd = rowTagEnd + 1;
        rowContent = '';
      } else {
        final closeIdx = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._rowClose, rowTagEnd);
        if (closeIdx == -1) {
          keepFrom = rowStart;
          break;
        }
        rowEnd = closeIdx + ExcelStreamReader._rowClose.length;
        rowContent = closeIdx > rowTagEnd + 1 ? utf8.decoder.convert(bytes, rowTagEnd + 1, closeIdx) : '';
      }
      pos = rowEnd;
      keepFrom = pos;

      final rowTag = utf8.decoder.convert(bytes, rowStart, rowTagEnd + 1);
      final rowIndexStr = ExcelStreamReader._extractAttrStream(rowTag, 'r');
      if (rowIndexStr == null) continue;
      final cells = rowContent.isEmpty ? const <ExcelCell>[] : _parseCells(rowContent);
      rows.add((row: ExcelRow(int.parse(rowIndexStr) - 1, cells), endOffset: _base + rowEnd));
    }
    if (end != -1 && pos >= end) {
      _done = true;
      _pending = Uint8List(0);
      return rows;
    }
    _keep(bytes, keepFrom);
    return rows;
  }

  void _keep(Uint8List bytes, int from) {
    _pending = from >= bytes.length ? Uint8List(0) : Uint8List.sublistView(bytes, from);
    _base += from;
  }

  static Uint8List _concat(Uint8List a, Uint8List b) {
    final out = Uint8List(a.length + b.length);
    out.setRange(0, a.length, a);
    out.setRange(a.length, out.length, b);
    return out;
  }
}

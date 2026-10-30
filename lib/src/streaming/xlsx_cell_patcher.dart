part of excel_community;

/// One cell [XlsxCellPatcher] rewrites: the 0-based [row] and [column] of the
/// sheet named [sheet]. A null or empty [text] leaves the cell empty.
class XlsxCellPatch {
  final String sheet;
  final int row;
  final int column;
  final String? text;

  const XlsxCellPatch({
    required this.sheet,
    required this.row,
    required this.column,
    this.text,
  });

  @override
  String toString() => '$sheet!${getCellId(column, row)}=${text ?? ''}';
}

/// Why [XlsxCellPatcher] left a patch out.
enum XlsxPatchMissReason {
  /// The workbook (in a package) or the sheet does not exist.
  noSheet,

  /// The sheet has no such row.
  noRow,

  /// The cell holds a formula: rewriting it could break a shared formula and
  /// the calculation chain, so it is corrected in the file.
  formula,
}

/// A patch [XlsxCellPatcher] did not write, and why.
class XlsxMissedPatch {
  final XlsxCellPatch patch;
  final XlsxPatchMissReason reason;

  const XlsxMissedPatch(this.patch, this.reason);

  @override
  String toString() => '$patch (${reason.name})';
}

/// Copies a workbook, or a zip that holds workbooks, with a few cells
/// rewritten, without decoding the workbook:
///
/// - every entry is copied byte for byte (its stored or deflated data as it
///   is, with its CRC and its time), except the XML of a sheet with patches;
/// - in that sheet only the rows with a patch are decoded, and in them only
///   the patched cells change: each becomes an inline string
///   (`t="inlineStr"`) that keeps the cell's `s` style; a missing cell is
///   inserted in column order;
/// - the shared strings, the styles, `workbook.xml` and the content types are
///   untouched.
///
/// The sheet is read and written one chunk at a time ([pause] is awaited
/// after each chunk and each entry, so a consumer can drain the sink), so a
/// sheet of any size costs one chunk and one row, not the sheet.
///
/// ```dart
/// final missed = await XlsxCellPatcher.writeWorkbook(bytes, sink, [
///   XlsxCellPatch(sheet: 'Units', row: 4, column: 1, text: 'U2'),
/// ]);
/// ```
class XlsxCellPatcher {
  XlsxCellPatcher._();

  /// [xlsx] with [patches] applied, as a new workbook into [sink]. Returns the
  /// patches it left out, with why: no sheet, no row, or a formula cell
  /// (nothing is written for them).
  static Future<List<XlsxMissedPatch>> writeWorkbook(
    List<int> xlsx,
    ByteSink sink,
    Iterable<XlsxCellPatch> patches, {
    Future<void> Function()? pause,
  }) async {
    final zip = ZipStreamWriter(sink);
    final missed = await _copyWorkbook(ZipDecoder().decodeBytes(xlsx), zip, patches, pause);
    zip.finish();
    return missed;
  }

  /// The zip [package] into [sink], every entry as it is except the workbooks
  /// named in [patches] (by their path in the zip), which are patched with
  /// [writeWorkbook]'s rules. Returns the patches that were not applied.
  static Future<List<XlsxMissedPatch>> writePackage(
    Archive package,
    ByteSink sink,
    Map<String, List<XlsxCellPatch>> patches, {
    Future<void> Function()? pause,
  }) async {
    final zip = ZipStreamWriter(sink);
    final missed = <XlsxMissedPatch>[];
    final pending = {...patches};
    for (final file in package.files) {
      final workbookPatches = file.isFile ? pending.remove(file.name) : null;
      if (workbookPatches == null) {
        copyEntry(zip, file);
      } else {
        zip.beginEntry(
          file.name,
          method: file.isCompressed ? ZipEntryMethod.deflate : ZipEntryMethod.stored,
          modified: file.lastModDateTime,
          large: true,
        );
        final workbook = ZipStreamWriter(_ZipEntrySink(zip));
        missed.addAll(await _copyWorkbook(ZipDecoder().decodeBytes(file.content), workbook, workbookPatches, pause));
        workbook.finish();
        zip.endEntry();
      }
      if (pause != null) await pause();
    }
    for (final unknown in pending.values) {
      missed.addAll([for (final patch in unknown) XlsxMissedPatch(patch, XlsxPatchMissReason.noSheet)]);
    }
    zip.finish();
    return missed;
  }

  /// Copies [file] of another archive into [zip] as it is: a zip entry's
  /// stored or deflated bytes with its CRC, without inflating them.
  static void copyEntry(ZipStreamWriter zip, ArchiveFile file) {
    final modified = file.lastModDateTime;
    if (!file.isFile) {
      zip.addDirectory(file.name, modified: modified);
      return;
    }
    final raw = file.rawContent;
    if (raw is ZipFile &&
        (raw.compressionMethod == CompressionType.deflate || raw.compressionMethod == CompressionType.none)) {
      final input = raw.getStream(decompress: false);
      final start = _entryStart[file] ??= input.position;
      input.setPosition(start);
      final bytes = input.toUint8List();
      input.setPosition(start);
      zip.addRawEntry(
        file.name,
        bytes,
        method: raw.compressionMethod == CompressionType.deflate ? ZipEntryMethod.deflate : ZipEntryMethod.stored,
        crc32: raw.crc32,
        uncompressedSize: raw.uncompressedSize,
        modified: modified,
      );
      return;
    }
    // Not from a zip (or another method): the content, deflated again.
    zip.addEntry(file.name, file.content, modified: modified);
  }

  static Future<List<XlsxMissedPatch>> _copyWorkbook(
    Archive archive,
    ZipStreamWriter zip,
    Iterable<XlsxCellPatch> patches,
    Future<void> Function()? pause,
  ) async {
    final targets = ExcelStreamReader._(archive)._sheetTargets;
    final missed = <XlsxMissedPatch>[];
    // Sheet part → row → column → patch.
    final bySheet = <String, Map<int, Map<int, XlsxCellPatch>>>{};
    for (final patch in patches) {
      final target = targets[patch.sheet];
      if (target == null) {
        missed.add(XlsxMissedPatch(patch, XlsxPatchMissReason.noSheet));
        continue;
      }
      ((bySheet[target] ??= {})[patch.row] ??= {})[patch.column] = patch;
    }

    for (final file in archive.files) {
      final rows = file.isFile ? bySheet[file.name] : null;
      if (rows == null) {
        copyEntry(zip, file);
      } else {
        // Not large: Excel asks to repair a workbook whose parts announce
        // zip64 without needing it.
        zip.beginEntry(file.name, modified: file.lastModDateTime);
        final patcher = _SheetPatcher(rows);
        for (final chunk in _sheetChunks(file)) {
          zip.write(patcher.feed(chunk));
          if (pause != null) await pause();
        }
        zip.write(patcher.finish());
        zip.endEntry();
        missed.addAll(patcher.refused);
        for (final MapEntry(key: row, value: cells) in rows.entries) {
          if (!patcher.patchedRows.contains(row)) {
            missed.addAll([for (final patch in cells.values) XlsxMissedPatch(patch, XlsxPatchMissReason.noRow)]);
          }
        }
      }
      if (pause != null) await pause();
    }
    return missed;
  }
}

/// The open entry of an outer zip as the sink of a workbook's own zip.
class _ZipEntrySink implements ByteSink {
  _ZipEntrySink(this._zip);

  final ZipStreamWriter _zip;

  @override
  void add(Uint8List bytes) => _zip.write(bytes);
}

/// Passes a sheet's XML through chunk by chunk, rewriting the rows of
/// [_rows]; a row cut by a chunk's end waits for the next chunk. Everything
/// outside `<sheetData>` and every other row goes out byte for byte.
class _SheetPatcher {
  _SheetPatcher(this._rows);

  final Map<int, Map<int, XlsxCellPatch>> _rows;

  /// The 0-based rows that were found (patched, or left as they are when
  /// every patch of the row was [refused]).
  final Set<int> patchedRows = {};

  /// The patches of found rows that were not written (a formula cell).
  final List<XlsxMissedPatch> refused = [];

  /// The 1-based number of the last row seen: a `<row>` without `r` is the
  /// one after it.
  int _lastRow = 0;

  static const int _before = 0;
  static const int _inside = 1;
  static const int _after = 2;

  static final int _keepTail = ExcelStreamReader._sheetDataClose.length - 1;

  int _state = _before;
  Uint8List _pending = Uint8List(0);

  /// The bytes to write for what came so far, [chunk] included.
  Uint8List feed(Uint8List chunk) {
    final bytes = _pending.isEmpty ? chunk : _RowScanner._concat(_pending, chunk);
    _pending = Uint8List(0);
    final out = BytesBuilder(copy: false);
    var pos = 0;

    void hold(int from) {
      out.add(Uint8List.sublistView(bytes, pos, from));
      _pending = Uint8List.fromList(Uint8List.sublistView(bytes, from));
    }

    while (true) {
      if (_state == _after) {
        out.add(Uint8List.sublistView(bytes, pos));
        return out.takeBytes();
      }
      if (_state == _before) {
        final open = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._sheetDataOpen, pos);
        if (open == -1) {
          hold(max(pos, bytes.length - ExcelStreamReader._sheetDataOpen.length + 1));
          return out.takeBytes();
        }
        final gt = ExcelStreamReader._indexOfByte(bytes, ExcelStreamReader._gt, open);
        if (gt == -1) {
          hold(open);
          return out.takeBytes();
        }
        _state = bytes[gt - 1] == ExcelStreamReader._slash ? _after : _inside;
        out.add(Uint8List.sublistView(bytes, pos, gt + 1));
        pos = gt + 1;
        continue;
      }

      final close = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._sheetDataClose, pos);
      final rowStart = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._rowOpen, pos);
      if (rowStart != -1 && (close == -1 || rowStart < close)) {
        final tagEnd = ExcelStreamReader._indexOfByte(bytes, ExcelStreamReader._gt, rowStart);
        if (tagEnd == -1) {
          hold(rowStart);
          return out.takeBytes();
        }
        final selfClosing = bytes[tagEnd - 1] == ExcelStreamReader._slash;
        var rowEnd = tagEnd + 1;
        var contentEnd = tagEnd + 1;
        if (!selfClosing) {
          final rowClose = ExcelStreamReader._indexOf(bytes, ExcelStreamReader._rowClose, tagEnd);
          if (rowClose == -1) {
            hold(rowStart);
            return out.takeBytes();
          }
          contentEnd = rowClose;
          rowEnd = rowClose + ExcelStreamReader._rowClose.length;
        }
        out.add(Uint8List.sublistView(bytes, pos, rowStart));
        final tag = utf8.decoder.convert(bytes, rowStart, tagEnd + 1);
        final r = ExcelStreamReader._extractAttrStream(tag, 'r');
        final row = (r == null ? null : int.tryParse(r)) ?? _lastRow + 1;
        _lastRow = row;
        final cells = _rows[row - 1];
        final content = cells == null || selfClosing ? '' : utf8.decoder.convert(bytes, tagEnd + 1, contentEnd);
        final patched = cells == null ? null : _patchRow(tag, content, row - 1, cells);
        if (cells != null) patchedRows.add(row - 1);
        out.add(patched == null ? Uint8List.sublistView(bytes, rowStart, rowEnd) : utf8.encode(patched));
        pos = rowEnd;
        continue;
      }
      if (close != -1) {
        out.add(Uint8List.sublistView(bytes, pos, close));
        pos = close;
        _state = _after;
        continue;
      }
      hold(max(pos, bytes.length - _keepTail));
      return out.takeBytes();
    }
  }

  /// What is still held back at the end of the sheet.
  Uint8List finish() {
    final rest = _pending;
    _pending = Uint8List(0);
    return rest;
  }

  /// The row [tag] + [content] with the cells of [patches] rewritten or
  /// inserted in column order. A row that gains a cell loses its `spans`
  /// hint, which may no longer hold. A cell with a formula is never
  /// rewritten ([refused]): it may be the master of a shared formula, and the
  /// calculation chain names it. Null when nothing in the row changes.
  String? _patchRow(String tag, String content, int row, Map<int, XlsxCellPatch> patches) {
    final remaining = patches.keys.toList()..sort();
    final out = StringBuffer();
    var inserted = false;
    var changed = false;
    var cursor = 0;
    var previousColumn = -1;
    var pos = 0;
    while (true) {
      final start = _cellStart(content, pos);
      if (start == -1) break;
      final gt = content.indexOf('>', start);
      if (gt == -1) break;
      final selfClosing = content.codeUnitAt(gt - 1) == 0x2F;
      final closeIdx = selfClosing ? -1 : content.indexOf('</c>', gt);
      if (!selfClosing && closeIdx == -1) break;
      final end = selfClosing ? gt + 1 : closeIdx + 4;
      final openingTag = content.substring(start, gt + 1);
      final reference = ExcelStreamReader._extractAttrStream(openingTag, 'r');
      final column = reference == null ? previousColumn + 1 : _cellCoordsFromCellId(reference).$2;
      previousColumn = column;

      out.write(content.substring(cursor, start));
      while (remaining.isNotEmpty && remaining.first < column) {
        out.write(_cell(row, remaining.removeAt(0), null, patches));
        inserted = true;
      }
      if (remaining.isNotEmpty && remaining.first == column) {
        remaining.removeAt(0);
        if (!selfClosing && content.substring(gt + 1, closeIdx).contains('<f')) {
          refused.add(XlsxMissedPatch(patches[column]!, XlsxPatchMissReason.formula));
          out.write(content.substring(start, end));
        } else {
          out.write(_cell(row, column, ExcelStreamReader._extractAttrStream(openingTag, 's'), patches));
          changed = true;
        }
      } else {
        out.write(content.substring(start, end));
      }
      cursor = end;
      pos = end;
    }
    final tail = content.substring(cursor);
    for (final column in remaining) {
      out.write(_cell(row, column, null, patches));
      inserted = true;
    }
    out.write(tail);
    if (!changed && !inserted) return null;

    var openTag = tag;
    if (inserted) openTag = openTag.replaceFirst(RegExp(r'\s+spans="[^"]*"'), '');
    if (openTag.endsWith('/>')) openTag = '${openTag.substring(0, openTag.length - 2).trimRight()}>';
    return '$openTag$out</row>';
  }

  /// The next `<c` that opens a cell element (not `<col…`), from [from].
  static int _cellStart(String content, int from) {
    var i = content.indexOf('<c', from);
    while (i != -1) {
      final next = i + 2 < content.length ? content.codeUnitAt(i + 2) : -1;
      if (next == 0x20 || next == 0x3E || next == 0x2F || next == 0x09 || next == 0x0A || next == 0x0D) return i;
      i = content.indexOf('<c', i + 2);
    }
    return -1;
  }

  static String _cell(int row, int column, String? style, Map<int, XlsxCellPatch> patches) {
    final text = patches[column]!.text;
    final reference = getCellId(column, row);
    final styleAttr = style == null ? '' : ' s="$style"';
    if (text == null || text.isEmpty) return '<c r="$reference"$styleAttr/>';
    // xml:space only when the text needs it: some readers look for a bare <t>.
    final space = text.trim() == text ? '' : ' xml:space="preserve"';
    return '<c r="$reference"$styleAttr t="inlineStr"><is><t$space>${_escape(text)}</t></is></c>';
  }

  /// [text] as XML character data; characters XML 1.0 cannot hold are left
  /// out.
  static String _escape(String text) {
    final out = StringBuffer();
    for (final rune in text.runes) {
      switch (rune) {
        case 0x26:
          out.write('&amp;');
        case 0x3C:
          out.write('&lt;');
        case 0x3E:
          out.write('&gt;');
        case 0x09 || 0x0A || 0x0D:
          out.writeCharCode(rune);
        default:
          if (rune >= 0x20 && rune != 0xFFFE && rune != 0xFFFF) out.writeCharCode(rune);
      }
    }
    return out.toString();
  }
}

part of excel_community;

/// A memory-efficient Excel reader that yields rows one at a time
/// without loading the entire sheet into memory.
///
/// Designed for large files (400k+ rows) in browser environments where
/// the DOM-based [Excel.decodeBytes] would run out of memory. A sheet is
/// scanned on its decompressed bytes, one `<row>` at a time: only the row
/// being read is ever decoded to a string, so memory is bounded by the
/// sheet's bytes plus one row, whatever the sheet size.
///
/// Cell values follow the cell's type and style: shared, inline and formula
/// strings become [TextCellValue], booleans [BoolCellValue], numbers
/// [IntCellValue] / [DoubleCellValue], and numbers whose cell style is a date
/// or time format become [DateCellValue], [DateTimeCellValue] or
/// [TimeCellValue], the same way [Excel.decodeBytes] reads them.
///
/// ```dart
/// final reader = ExcelStreamReader.fromBytes(xlsxBytes);
/// print(reader.sheetNames); // ['Sheet1', 'Sheet2']
///
/// for (final row in reader.readSheet('Sheet1')) {
///   print('Row ${row.index}: ${row.cells.length} cells');
///   // Process row — it is discarded when you move to the next
/// }
///
/// // On a UI thread, let the event loop breathe between slices:
/// await for (final row in reader.readSheetAsync('Sheet1', pause: yieldToUi)) { … }
/// ```
class ExcelStreamReader {
  /// Bytes consumed between two [readSheetAsync] pauses.
  static const int defaultSliceSize = 64 * 1024;

  final Archive _archive;
  final Map<String, String> _sheetTargets = {}; // name → xl/worksheets/sheetN.xml
  List<String>? _sharedStringsList;

  /// Number format per `cellXfs` index, from `xl/styles.xml`; null when the
  /// style has no (known) number format.
  List<NumFormat?>? _xfFormats;

  ExcelStreamReader._(this._archive) {
    _parseWorkbook();
  }

  /// Create a reader from XLSX bytes.
  factory ExcelStreamReader.fromBytes(List<int> bytes) {
    final archive = ZipDecoder().decodeBytes(bytes);
    return ExcelStreamReader._(archive);
  }

  /// Available sheet names in the workbook.
  List<String> get sheetNames => _sheetTargets.keys.toList();

  /// The number of rows the sheet's `<dimension ref="A1:N279520"/>` declares,
  /// header row included; null when the sheet carries no dimension. Cheap:
  /// only the head of the sheet is looked at.
  int? rowCount(String sheetName) {
    final bytes = _sheetFile(sheetName).content;
    final head = utf8.decoder.convert(bytes, 0, min(bytes.length, 4096));
    final match =
        RegExp(r'<dimension\s+ref="[A-Z]+(\d+)(?::[A-Z]+(\d+))?"').firstMatch(head);
    if (match == null) return null;
    return int.tryParse(match.group(2) ?? match.group(1)!);
  }

  /// Read rows from a sheet as a lazy iterable.
  /// Each [ExcelRow] is yielded one at a time for memory efficiency.
  Iterable<ExcelRow> readSheet(String sheetName) sync* {
    final file = _sheetFile(sheetName);
    _loadSharedStrings();
    for (final scanned in _scanRows(file.content)) {
      yield scanned.row;
    }
  }

  /// [readSheet] for a UI thread: the same rows, with [pause] awaited after
  /// every [sliceSize] bytes of sheet data, so a progress dialog can repaint
  /// between slices (`() => Future.delayed(Duration.zero)` is enough on the
  /// web). Without [pause] it is a plain stream of the rows.
  Stream<ExcelRow> readSheetAsync(
    String sheetName, {
    Future<void> Function()? pause,
    int sliceSize = defaultSliceSize,
  }) async* {
    // Inflating the sheet and loading the shared strings are the two
    // synchronous blocks before the first row; give the caller a turn after
    // each so a progress dialog can paint.
    final file = _sheetFile(sheetName);
    if (pause != null) await pause();
    _loadSharedStrings();
    if (pause != null) await pause();
    var nextPause = sliceSize;
    for (final scanned in _scanRows(file.content)) {
      yield scanned.row;
      if (pause != null && scanned.endOffset >= nextPause) {
        nextPause = scanned.endOffset + sliceSize;
        await pause();
      }
    }
  }

  ArchiveFile _sheetFile(String sheetName) {
    final target = _sheetTargets[sheetName];
    if (target == null) {
      throw ArgumentError('Sheet "$sheetName" not found. '
          'Available: ${sheetNames.join(', ')}');
    }
    final file = _archive.findFile(target);
    if (file == null) {
      throw StateError('Worksheet file "$target" not found in archive');
    }
    file.decompress();
    return file;
  }

  void _parseWorkbook() {
    // Parse workbook relationships to get sheet targets
    final relsFile = _archive.findFile('xl/_rels/workbook.xml.rels');
    if (relsFile == null) return;
    relsFile.decompress();
    final relsDoc = XmlDocument.parse(utf8.decode(relsFile.content));

    final targetMap = <String, String>{}; // rId → target
    for (final rel in relsDoc.findAllElements('Relationship')) {
      final id = rel.getAttribute('Id');
      final target = rel.getAttribute('Target');
      final type = rel.getAttribute('Type');
      if (id != null &&
          target != null &&
          type ==
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet') {
        targetMap[id] = target.startsWith('/') ? target.substring(1) : 'xl/$target';
      }
    }

    // Parse workbook.xml for sheet names
    final workbookFile = _archive.findFile('xl/workbook.xml');
    if (workbookFile == null) return;
    workbookFile.decompress();
    final workbookDoc = XmlDocument.parse(utf8.decode(workbookFile.content));

    for (final sheet in workbookDoc.findAllElements('sheet')) {
      final name = sheet.getAttribute('name');
      final rId = sheet.getAttribute('r:id');
      if (name != null && rId != null && targetMap.containsKey(rId)) {
        _sheetTargets[name] = targetMap[rId]!;
      }
    }
  }

  /// Number formats by `cellXfs` index, the way the DOM parser resolves them
  /// (`_parseStyles`): custom `numFmt`s registered on a [NumFormatMaintainer]
  /// that already knows the built-in ids, then one lookup per `xf`.
  void _loadStyles() {
    if (_xfFormats != null) return;
    final formats = <NumFormat?>[];
    _xfFormats = formats;

    final stylesFile = _archive.findFile('xl/styles.xml');
    if (stylesFile == null) return;
    stylesFile.decompress();
    final document = XmlDocument.parse(utf8.decode(stylesFile.content));

    final maintainer = NumFormatMaintainer();
    for (final numFmts in document.findAllElements('numFmts')) {
      for (final node in numFmts.findAllElements('numFmt')) {
        final numFmtId = int.tryParse(node.getAttribute('numFmtId') ?? '');
        final formatCode = node.getAttribute('formatCode');
        if (numFmtId == null || formatCode == null) continue;
        if (numFmtId >= 164 && maintainer.getByNumFmtId(numFmtId) == null) {
          maintainer.add(numFmtId, NumFormat.custom(formatCode: formatCode));
        }
      }
    }
    for (final cellXfs in document.findAllElements('cellXfs')) {
      for (final xf in cellXfs.findAllElements('xf')) {
        final numFmtId = int.tryParse(xf.getAttribute('numFmtId') ?? '0') ?? 0;
        formats.add(maintainer.getByNumFmtId(numFmtId));
      }
    }
  }

  void _loadSharedStrings() {
    _loadStyles();
    if (_sharedStringsList != null) return;
    _sharedStringsList = [];

    // Try common locations
    ArchiveFile? ssFile;
    for (final path in [
      'xl/sharedStrings.xml',
      'xl/SharedStrings.xml',
    ]) {
      ssFile = _archive.findFile(path);
      if (ssFile != null) break;
    }

    if (ssFile == null) return;
    ssFile.decompress();
    final bytes = ssFile.content;
    final list = _sharedStringsList!;

    // Scan <si>…</si> on the bytes, like the rows: a 10 MB table of a few
    // hundred thousand strings costs one small string per item instead of
    // a DOM of the whole part. Every <si>, empty or not, is one index.
    var pos = 0;
    while (true) {
      final siStart = _indexOf(bytes, _siOpen, pos);
      if (siStart == -1) break;
      final afterTag = siStart + _siOpen.length;
      if (afterTag >= bytes.length) break;
      if (bytes[afterTag] == _slash) {
        list.add(''); // <si/>
        pos = afterTag + 2;
        continue;
      }
      final siEnd = _indexOf(bytes, _siClose, afterTag);
      if (siEnd == -1) break;
      list.add(_sharedItemText(utf8.decoder.convert(bytes, afterTag + 1, siEnd)));
      pos = siEnd + _siClose.length;
    }
  }

  /// The text of one shared string item: its `<t>` runs concatenated,
  /// phonetic runs (`<rPh>`) left out, entities decoded.
  static String _sharedItemText(String si) {
    final withoutPhonetic =
        si.contains('<rPh') ? si.replaceAll(_phoneticRun, '') : si;
    return _inlineText(withoutPhonetic);
  }

  static final RegExp _phoneticRun = RegExp(r'<rPh[\s\S]*?</rPh>');

  // ── Byte-level row scanning ──────────────────────────────────────

  static final List<int> _rowOpen = utf8.encode('<row');
  static final List<int> _rowClose = utf8.encode('</row>');
  static final List<int> _siOpen = utf8.encode('<si');
  static final List<int> _siClose = utf8.encode('</si>');
  static final List<int> _sheetDataOpen = utf8.encode('<sheetData');
  static final List<int> _sheetDataClose = utf8.encode('</sheetData>');
  static const int _gt = 0x3E; // '>'
  static const int _slash = 0x2F; // '/'

  /// Walks the sheet's bytes `<row>` by `<row>`. The markup is ASCII, and a
  /// UTF-8 multi-byte sequence never contains an ASCII byte, so searching
  /// the bytes is safe; only one row at a time is decoded to a string.
  /// [endOffset] is the byte offset just past the row, for pausing.
  Iterable<({ExcelRow row, int endOffset})> _scanRows(List<int> bytes) sync* {
    final sheetDataStart = _indexOf(bytes, _sheetDataOpen, 0);
    if (sheetDataStart == -1) return;
    final afterOpen = _indexOfByte(bytes, _gt, sheetDataStart);
    if (afterOpen == -1) return;
    if (bytes[afterOpen - 1] == _slash) return; // <sheetData/>
    var end = _indexOf(bytes, _sheetDataClose, afterOpen);
    if (end == -1) end = bytes.length;

    var pos = afterOpen + 1;
    while (pos < end) {
      final rowStart = _indexOf(bytes, _rowOpen, pos);
      if (rowStart == -1 || rowStart >= end) break;
      final rowTagEnd = _indexOfByte(bytes, _gt, rowStart);
      if (rowTagEnd == -1) break;

      final isSelfClosing = bytes[rowTagEnd - 1] == _slash;
      final rowTag = utf8.decoder.convert(bytes, rowStart, rowTagEnd + 1);
      final rowIndexStr = _extractAttrStream(rowTag, 'r');

      int rowEnd;
      String rowContent;
      if (isSelfClosing) {
        rowEnd = rowTagEnd + 1;
        rowContent = '';
      } else {
        final closeIdx = _indexOf(bytes, _rowClose, rowTagEnd);
        if (closeIdx == -1) break;
        rowEnd = closeIdx + _rowClose.length;
        rowContent = closeIdx > rowTagEnd + 1
            ? utf8.decoder.convert(bytes, rowTagEnd + 1, closeIdx)
            : '';
      }
      pos = rowEnd;

      if (rowIndexStr == null) continue;
      final rowIdx = int.parse(rowIndexStr) - 1;
      final cells =
          rowContent.isEmpty ? const <ExcelCell>[] : _parseCells(rowContent);
      yield (row: ExcelRow(rowIdx, cells), endOffset: rowEnd);
    }
  }

  static int _indexOf(List<int> bytes, List<int> needle, int from) {
    final first = needle[0];
    final last = bytes.length - needle.length;
    for (var i = from; i <= last; i++) {
      if (bytes[i] != first) continue;
      var j = 1;
      while (j < needle.length && bytes[i + j] == needle[j]) {
        j++;
      }
      if (j == needle.length) return i;
    }
    return -1;
  }

  static int _indexOfByte(List<int> bytes, int byte, int from) {
    for (var i = from; i < bytes.length; i++) {
      if (bytes[i] == byte) return i;
    }
    return -1;
  }

  // ── Cell parsing ─────────────────────────────────────────────────

  List<ExcelCell> _parseCells(String content) {
    final cells = <ExcelCell>[];
    int pos = 0;
    final len = content.length;

    while (pos < len) {
      final cellStart = content.indexOf('<c ', pos);
      if (cellStart == -1) break;

      // Find cell end
      int cellEnd;
      final selfCloseCheck = content.indexOf('/>', cellStart);
      final childCloseCheck = content.indexOf('</c>', cellStart);

      if (childCloseCheck == -1 && selfCloseCheck == -1) break;

      bool isSelfClosing;
      if (childCloseCheck == -1) {
        isSelfClosing = true;
        cellEnd = selfCloseCheck + 2;
      } else if (selfCloseCheck == -1) {
        isSelfClosing = false;
        cellEnd = childCloseCheck + 4;
      } else {
        final firstGt = content.indexOf('>', cellStart);
        if (firstGt != -1 && selfCloseCheck == firstGt - 1) {
          isSelfClosing = true;
          cellEnd = selfCloseCheck + 2;
        } else {
          isSelfClosing = false;
          cellEnd = childCloseCheck + 4;
        }
      }

      final cellXml = content.substring(cellStart, cellEnd);
      final tagEnd = cellXml.indexOf('>');
      final openingTag = cellXml.substring(0, tagEnd + 1);

      final rAttr = _extractAttrStream(openingTag, 'r');
      if (rAttr == null) {
        pos = cellEnd;
        continue;
      }

      final coords = _cellCoordsFromCellId(rAttr);
      final columnIndex = coords.$2;
      final tAttr = _extractAttrStream(openingTag, 't');
      final sAttr = _extractAttrStream(openingTag, 's');

      CellValue? value;

      if (!isSelfClosing) {
        final childContent =
            cellXml.substring(tagEnd + 1, cellXml.length - 4);

        switch (tAttr) {
          case 's': // shared string
            final vVal = _extractElementTextStream(childContent, 'v');
            if (vVal != null && _sharedStringsList != null) {
              final idx = int.tryParse(vVal.trim());
              if (idx != null && idx < _sharedStringsList!.length) {
                value = TextCellValue(_sharedStringsList![idx]);
              }
            }
            break;
          case 'str': // formula result as text
            final vVal = _extractElementTextStream(childContent, 'v');
            if (vVal != null) value = TextCellValue(_unescapeXml(vVal));
            break;
          case 'inlineStr':
            value = TextCellValue(_inlineText(childContent));
            break;
          case 'b': // boolean
            final vVal = _extractElementTextStream(childContent, 'v');
            value = BoolCellValue(vVal == '1');
            break;
          case 'e': // error
            break;
          default: // number, a date when the style says so
            final vVal = _extractElementTextStream(childContent, 'v');
            if (vVal != null) {
              value = _numericValue(vVal, sAttr);
            }
            break;
        }
      }

      if (value != null) {
        cells.add(ExcelCell(columnIndex, value));
      }

      pos = cellEnd;
    }

    return cells;
  }

  CellValue? _numericValue(String vVal, String? sAttr) {
    final styleIndex = sAttr == null ? null : int.tryParse(sAttr);
    final formats = _xfFormats;
    if (styleIndex != null && formats != null && styleIndex < formats.length) {
      final format = formats[styleIndex];
      if (format is DateTimeNumFormat) {
        return format.read(vVal);
      }
    }
    final intVal = int.tryParse(vVal);
    if (intVal != null) return IntCellValue(intVal);
    final doubleVal = double.tryParse(vVal);
    if (doubleVal != null) return DoubleCellValue(doubleVal);
    return null;
  }

  /// The text of an `<is>` inline string: every `<t>` concatenated (rich
  /// text has several runs), entities decoded.
  static String _inlineText(String childContent) {
    final buf = StringBuffer();
    for (final match in RegExp(r'<t(?:\s[^>]*)?>(.*?)</t>', dotAll: true)
        .allMatches(childContent)) {
      buf.write(_unescapeXml(match.group(1)!));
    }
    return buf.toString();
  }

  static String _unescapeXml(String text) {
    if (!text.contains('&')) return text;
    return text
        .replaceAll('&lt;', '<')
        .replaceAll('&gt;', '>')
        .replaceAll('&quot;', '"')
        .replaceAll('&apos;', "'")
        .replaceAllMapped(RegExp(r'&#x([0-9A-Fa-f]+);'),
            (m) => String.fromCharCode(int.parse(m.group(1)!, radix: 16)))
        .replaceAllMapped(RegExp(r'&#(\d+);'),
            (m) => String.fromCharCode(int.parse(m.group(1)!)))
        .replaceAll('&amp;', '&');
  }

  /// Extract an attribute value from an XML tag string.
  static String? _extractAttrStream(String tag, String name) {
    // Search for ' name="value"' pattern (space before name prevents
    // partial matches, e.g. 'spans' matching 's')
    final search = ' $name="';
    final idx = tag.indexOf(search);
    if (idx == -1) return null;
    final start = idx + search.length;
    final end = tag.indexOf('"', start);
    if (end == -1) return null;
    return tag.substring(start, end);
  }

  /// Extract text content of a simple XML element.
  static String? _extractElementTextStream(String content, String tag) {
    final openTag = '<$tag>';
    final closeTag = '</$tag>';
    final openIdx = content.indexOf(openTag);
    if (openIdx == -1) return null;
    final start = openIdx + openTag.length;
    final end = content.indexOf(closeTag, start);
    if (end == -1) return null;
    return content.substring(start, end);
  }
}

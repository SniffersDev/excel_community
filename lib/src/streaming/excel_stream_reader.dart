part of excel_community;

/// A memory-efficient Excel reader that yields rows one at a time
/// without loading the entire sheet into memory.
///
/// Designed for large files (400k+ rows) in browser environments where
/// the DOM-based [Excel.decodeBytes] would run out of memory.
///
/// ```dart
/// final reader = ExcelStreamReader.fromBytes(xlsxBytes);
/// print(reader.sheetNames); // ['Sheet1', 'Sheet2']
///
/// for (final row in reader.readSheet('Sheet1')) {
///   print('Row ${row.index}: ${row.cells.length} cells');
///   // Process row — it is discarded when you move to the next
/// }
/// ```
class ExcelStreamReader {
  final Archive _archive;
  final Map<String, String> _sheetTargets = {}; // name → xl/worksheets/sheetN.xml
  List<String>? _sharedStringsList;

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

  /// Read rows from a sheet as a lazy iterable.
  /// Each [ExcelRow] is yielded one at a time for memory efficiency.
  Iterable<ExcelRow> readSheet(String sheetName) sync* {
    final target = _sheetTargets[sheetName];
    if (target == null) {
      throw ArgumentError('Sheet "$sheetName" not found. '
          'Available: ${sheetNames.join(', ')}');
    }

    // Ensure shared strings are loaded
    _loadSharedStrings();

    // Extract worksheet XML
    final file = _archive.findFile(target);
    if (file == null) {
      throw StateError('Worksheet file "$target" not found in archive');
    }
    file.decompress();
    final rawXml = utf8.decode(file.content);

    // Find <sheetData> section
    final sheetDataOpenIdx = rawXml.indexOf('<sheetData');
    if (sheetDataOpenIdx == -1) return;

    final afterOpen = rawXml.indexOf('>', sheetDataOpenIdx);
    if (afterOpen == -1) return;

    // Check self-closing <sheetData/>
    if (rawXml[afterOpen - 1] == '/') return;

    final sheetDataCloseIdx = rawXml.indexOf('</sheetData>', afterOpen);
    if (sheetDataCloseIdx == -1) return;

    final sheetDataContent = rawXml.substring(afterOpen + 1, sheetDataCloseIdx);

    // Parse rows from the raw XML string
    yield* _parseRows(sheetDataContent);
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

  void _loadSharedStrings() {
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
    final ssDoc = XmlDocument.parse(utf8.decode(ssFile.content));

    for (final si in ssDoc.findAllElements('si')) {
      final buf = StringBuffer();
      for (final t in si.findAllElements('t')) {
        // Skip phonetic run text
        if (t.parentElement != null &&
            t.parentElement!.name.local == 'rPh') {
          continue;
        }
        buf.write(t.innerText);
      }
      _sharedStringsList!.add(buf.toString());
    }
  }

  Iterable<ExcelRow> _parseRows(String content) sync* {
    int pos = 0;
    final len = content.length;

    while (pos < len) {
      final rowStart = content.indexOf('<row', pos);
      if (rowStart == -1) break;

      final rowTagEnd = content.indexOf('>', rowStart);
      if (rowTagEnd == -1) break;

      final isSelfClosing = content[rowTagEnd - 1] == '/';
      final rowTag = content.substring(rowStart, rowTagEnd + 1);
      final rowIndexStr = _extractAttrStream(rowTag, 'r');

      int rowEnd;
      String rowContent;

      if (isSelfClosing) {
        rowEnd = rowTagEnd + 1;
        rowContent = '';
      } else {
        final closeIdx = content.indexOf('</row>', rowTagEnd);
        if (closeIdx == -1) break;
        rowEnd = closeIdx + 6;
        rowContent = content.substring(rowTagEnd + 1, closeIdx);
      }

      if (rowIndexStr == null) {
        pos = rowEnd;
        continue;
      }

      final rowIdx = int.parse(rowIndexStr) - 1;

      if (rowContent.isNotEmpty) {
        final cells = _parseCells(rowContent);
        yield ExcelRow(rowIdx, cells);
      } else {
        yield ExcelRow(rowIdx, const []);
      }

      pos = rowEnd;
    }
  }

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
          case 'b': // boolean
            final vVal = _extractElementTextStream(childContent, 'v');
            value = BoolCellValue(vVal == '1');
            break;
          default: // number
            final vVal = _extractElementTextStream(childContent, 'v');
            if (vVal != null) {
              final intVal = int.tryParse(vVal);
              if (intVal != null) {
                value = IntCellValue(intVal);
              } else {
                final doubleVal = double.tryParse(vVal);
                if (doubleVal != null) {
                  value = DoubleCellValue(doubleVal);
                }
              }
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

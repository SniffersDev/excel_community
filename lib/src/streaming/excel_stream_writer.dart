part of excel_community;

/// A memory-efficient Excel writer that writes rows directly to XML
/// StringBuffers without building Sheet/Data DOM objects.
///
/// Designed for large files (400k+ rows) in browser environments where
/// the DOM-based [Excel] class would run out of memory.
///
/// Supports multiple sheets, mixed cell types, shared strings, and
/// optional cell styling (fonts, fills, borders, alignment).
/// Does NOT support charts, merged cells, or formulas.
///
/// ```dart
/// final headerStyle = CellStyle(
///   bold: true,
///   fontColorHex: ExcelColor.white,
///   backgroundColorHex: ExcelColor.blue,
/// );
///
/// final writer = ExcelStreamWriter(sheetName: 'Sales');
/// writer.addHeaderRow(['Month', 'Revenue'], headerStyle: headerStyle);
/// writer.addRow([TextCellValue('Jan'), IntCellValue(1000)]);
///
/// final bytes = writer.encode();
/// ```
class ExcelStreamWriter {
  /// Shared strings: value → index (global across all sheets).
  final Map<String, int> _sharedStrings = {};
  int _sharedStringCount = 0;

  /// Per-sheet data.
  final List<_StreamSheet> _sheets = [];
  late _StreamSheet _activeSheet;

  bool _encoded = false;

  // ── Style registries ──────────────────────────────────────────────
  // Index 0 in each list is the default (unstyled) entry.

  /// Unique font definitions. Index 0 = default Calibri 11.
  final List<_FontStyle> _fonts = [
    _FontStyle(fontSize: 11, fontFamily: 'Calibri'),
  ];

  /// Unique fill patterns. Index 0 = none, index 1 = gray125 (required by Excel).
  final List<String> _fills = ['none', 'gray125'];

  /// Unique border sets. Index 0 = empty borders.
  final List<_BorderSet> _borders = [
    _BorderSet(
      leftBorder: Border(),
      rightBorder: Border(),
      topBorder: Border(),
      bottomBorder: Border(),
      diagonalBorder: Border(),
      diagonalBorderUp: false,
      diagonalBorderDown: false,
    ),
  ];

  /// Unique cellXf records (font+fill+border+alignment combo).
  /// Each entry is a tuple: (fontId, fillId, borderId, CellStyle?).
  /// Index 0 = default unstyled xf.
  final List<_StreamXf> _xfs = [
    _StreamXf(fontId: 0, fillId: 0, borderId: 0, style: null),
  ];

  /// Cache: CellStyle → xfId for deduplication.
  final Map<CellStyle, int> _styleToXfId = {};

  /// Create a new streaming writer. The first sheet is created automatically.
  ExcelStreamWriter({String sheetName = 'Sheet1'}) {
    _activeSheet = _StreamSheet(sheetName);
    _sheets.add(_activeSheet);
  }

  /// Add a new sheet and make it the active target for [addRow].
  ExcelStreamWriter addSheet(String name) {
    if (_encoded) throw StateError('Cannot add sheets after encode()');
    _activeSheet = _StreamSheet(name);
    _sheets.add(_activeSheet);
    return this;
  }

  /// Add a header row of text values to the active sheet.
  ///
  /// If [headerStyle] is provided, every cell in the header row
  /// will use that style (e.g. bold, colored background).
  void addHeaderRow(List<String> headers, {CellStyle? headerStyle}) {
    final values = headers.map((h) => TextCellValue(h)).toList();
    final styles = headerStyle != null
        ? List<CellStyle?>.filled(headers.length, headerStyle)
        : null;
    addRow(values, styles: styles);
  }

  /// Add a data row to the active sheet.
  ///
  /// [styles] is an optional parallel list of [CellStyle] for each cell.
  /// Pass `null` for unstyled cells within the list.
  void addRow(List<CellValue?> values, {List<CellStyle?>? styles}) {
    if (_encoded) throw StateError('Cannot add rows after encode()');

    final rowIndex = _activeSheet.rowCount;
    final buffer = _activeSheet.buffer;

    buffer.write('<row r="');
    buffer.write(rowIndex + 1);
    buffer.write('">');

    for (var col = 0; col < values.length; col++) {
      final value = values[col];
      if (value == null) continue;
      final style =
          (styles != null && col < styles.length) ? styles[col] : null;
      _writeCell(buffer, col, rowIndex, value, style);
    }

    buffer.write('</row>');
    _activeSheet.rowCount++;
  }

  // ── Cell writing ─────────────────────────────────────────────────

  void _writeCell(
      StringBuffer buffer, int col, int row, CellValue value,
      [CellStyle? style]) {
    final cellRef = getCellId(col, row);

    buffer.write('<c r="');
    buffer.write(cellRef);
    buffer.write('"');

    if (value is TextCellValue) {
      buffer.write(' t="s"');
    } else if (value is BoolCellValue) {
      buffer.write(' t="b"');
    }

    // Apply style index if provided
    if (style != null) {
      final xfId = _resolveStyleId(style);
      buffer.write(' s="');
      buffer.write(xfId);
      buffer.write('"');
    }

    buffer.write('><v>');

    switch (value) {
      case TextCellValue():
        final text = value.toString();
        var idx = _sharedStrings[text];
        if (idx == null) {
          idx = _sharedStringCount++;
          _sharedStrings[text] = idx;
        }
        buffer.write(idx);
        break;
      case IntCellValue():
        buffer.write(value.write(null));
        break;
      case DoubleCellValue():
        buffer.write(value.write(null));
        break;
      case BoolCellValue():
        buffer.write(value.write(null));
        break;
      case DateCellValue():
        buffer.write(value.write(null));
        break;
      case TimeCellValue():
        buffer.write(value.write(null));
        break;
      case DateTimeCellValue():
        buffer.write(value.write(null));
        break;
      case FormulaCellValue():
        break;
    }

    buffer.write('</v></c>');
  }

  // ── Style resolution ────────────────────────────────────────────

  /// Returns the xfId for a [CellStyle], registering new font/fill/border
  /// entries as needed. Uses a cache for deduplication.
  int _resolveStyleId(CellStyle style) {
    final cached = _styleToXfId[style];
    if (cached != null) return cached;

    // -- Font --
    final font = _FontStyle(
      bold: style.isBold,
      italic: style.isItalic,
      fontColorHex: style.fontColor,
      underline: style.underline,
      fontSize: style.fontSize,
      fontFamily: style.fontFamily,
    );
    int fontId = _fonts.indexOf(font);
    if (fontId == -1) {
      fontId = _fonts.length;
      _fonts.add(font);
    }

    // -- Fill --
    final bgHex = style.backgroundColor.colorHex;
    int fillId;
    if (bgHex == 'none') {
      fillId = 0;
    } else {
      fillId = _fills.indexOf(bgHex);
      if (fillId == -1) {
        fillId = _fills.length;
        _fills.add(bgHex);
      }
    }

    // -- Border --
    final borderSet = _BorderSet(
      leftBorder: style.leftBorder,
      rightBorder: style.rightBorder,
      topBorder: style.topBorder,
      bottomBorder: style.bottomBorder,
      diagonalBorder: style.diagonalBorder,
      diagonalBorderUp: style.diagonalBorderUp,
      diagonalBorderDown: style.diagonalBorderDown,
    );
    int borderId = _borders.indexOf(borderSet);
    if (borderId == -1) {
      borderId = _borders.length;
      _borders.add(borderSet);
    }

    // -- XF record --
    final xf = _StreamXf(
        fontId: fontId, fillId: fillId, borderId: borderId, style: style);
    int xfId = _xfs.indexOf(xf);
    if (xfId == -1) {
      xfId = _xfs.length;
      _xfs.add(xf);
    }

    _styleToXfId[style] = xfId;
    return xfId;
  }

  // ── Encoding ─────────────────────────────────────────────────────

  /// Encode all sheets into XLSX bytes.
  List<int> encode() {
    if (_encoded) throw StateError('encode() can only be called once');
    _encoded = true;

    final archive = Archive();

    // 1. [Content_Types].xml
    archive.addFile(_makeArchiveFile(
        '[Content_Types].xml', _buildContentTypes()));

    // 2. _rels/.rels
    archive.addFile(_makeArchiveFile('_rels/.rels', _buildRootRels()));

    // 3. xl/_rels/workbook.xml.rels
    archive.addFile(_makeArchiveFile(
        'xl/_rels/workbook.xml.rels', _buildWorkbookRels()));

    // 4. xl/workbook.xml
    archive.addFile(
        _makeArchiveFile('xl/workbook.xml', _buildWorkbook()));

    // 5. xl/styles.xml
    archive.addFile(
        _makeArchiveFile('xl/styles.xml', _buildStyles()));

    // 6. xl/sharedStrings.xml
    archive.addFile(_makeArchiveFile(
        'xl/sharedStrings.xml', _buildSharedStrings()));

    // 7. xl/worksheets/sheetN.xml for each sheet
    for (var i = 0; i < _sheets.length; i++) {
      archive.addFile(_makeArchiveFile(
          'xl/worksheets/sheet${i + 1}.xml',
          _buildWorksheet(_sheets[i])));
    }

    return ZipEncoder().encode(archive);
  }

  /// Web convenience: encode + trigger browser download.
  List<int>? save({String fileName = 'ExcelStream.xlsx'}) {
    final bytes = encode();
    return helper.SavingHelper.saveFile(bytes, fileName);
  }

  ArchiveFile _makeArchiveFile(String name, String content) {
    final bytes = utf8.encode(content);
    return ArchiveFile(name, bytes.length, bytes)
      ..compression = _noCompression.contains(name)
          ? CompressionType.none
          : CompressionType.deflate;
  }

  // ── XML builders ──────────────────────────────────────────────────

  String _buildContentTypes() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
    buf.write('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    buf.write('<Default Extension="xml" ContentType="application/xml"/>');
    buf.write('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>');
    buf.write('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
    buf.write('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
    for (var i = 0; i < _sheets.length; i++) {
      buf.write('<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
    }
    buf.write('</Types>');
    return buf.toString();
  }

  String _buildRootRels() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>';
  }

  String _buildWorkbookRels() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
    for (var i = 0; i < _sheets.length; i++) {
      buf.write('<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>');
    }
    final stylesRid = _sheets.length + 1;
    final sharedStringsRid = _sheets.length + 2;
    buf.write('<Relationship Id="rId$stylesRid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');
    buf.write('<Relationship Id="rId$sharedStringsRid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>');
    buf.write('</Relationships>');
    return buf.toString();
  }

  String _buildWorkbook() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    buf.write('<sheets>');
    for (var i = 0; i < _sheets.length; i++) {
      buf.write('<sheet name="${_escapeXmlStream(_sheets[i].name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>');
    }
    buf.write('</sheets>');
    buf.write('</workbook>');
    return buf.toString();
  }

  // ── Full styles.xml builder ──────────────────────────────────────

  String _buildStyles() {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

    // Fonts
    buf.write('<fonts count="${_fonts.length}">');
    for (final font in _fonts) {
      buf.write('<font>');
      if (font.isBold) buf.write('<b/>');
      if (font.isItalic) buf.write('<i/>');
      if (font.underline == Underline.Single) {
        buf.write('<u/>');
      } else if (font.underline == Underline.Double) {
        buf.write('<u val="double"/>');
      }
      if (font.fontSize != null) {
        buf.write('<sz val="${font.fontSize}"/>');
      }
      if (font.fontColor.colorHex != 'FF000000' &&
          font.fontColor.colorHex != 'none') {
        buf.write('<color rgb="${font.fontColor.colorHex}"/>');
      }
      if (font.fontFamily != null && font.fontFamily!.isNotEmpty) {
        buf.write('<name val="${_escapeXmlStream(font.fontFamily!)}"/>');
      }
      buf.write('</font>');
    }
    buf.write('</fonts>');

    // Fills
    buf.write('<fills count="${_fills.length}">');
    for (final fill in _fills) {
      if (fill == 'none' || fill == 'gray125' || fill == 'lightGray') {
        buf.write('<fill><patternFill patternType="$fill"/></fill>');
      } else {
        // Solid color fill
        buf.write('<fill><patternFill patternType="solid">');
        buf.write('<fgColor rgb="$fill"/>');
        buf.write('<bgColor rgb="$fill"/>');
        buf.write('</patternFill></fill>');
      }
    }
    buf.write('</fills>');

    // Borders
    buf.write('<borders count="${_borders.length}">');
    for (final bs in _borders) {
      buf.write('<border');
      if (bs.diagonalBorderDown) buf.write(' diagonalDown="1"');
      if (bs.diagonalBorderUp) buf.write(' diagonalUp="1"');
      buf.write('>');

      _writeBorderSide(buf, 'left', bs.leftBorder);
      _writeBorderSide(buf, 'right', bs.rightBorder);
      _writeBorderSide(buf, 'top', bs.topBorder);
      _writeBorderSide(buf, 'bottom', bs.bottomBorder);
      _writeBorderSide(buf, 'diagonal', bs.diagonalBorder);

      buf.write('</border>');
    }
    buf.write('</borders>');

    // CellStyleXfs (base styles)
    buf.write('<cellStyleXfs count="1">');
    buf.write('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
    buf.write('</cellStyleXfs>');

    // CellXfs (actual cell format records)
    buf.write('<cellXfs count="${_xfs.length}">');
    for (final xf in _xfs) {
      buf.write('<xf');
      buf.write(' numFmtId="0"');
      buf.write(' fontId="${xf.fontId}"');
      buf.write(' fillId="${xf.fillId}"');
      buf.write(' borderId="${xf.borderId}"');
      buf.write(' xfId="0"');
      if (xf.fontId != 0) buf.write(' applyFont="1"');
      if (xf.fillId != 0) buf.write(' applyFill="1"');
      if (xf.borderId != 0) buf.write(' applyBorder="1"');

      final style = xf.style;
      final hasAlignment = style != null &&
          (style.horizontalAlignment != HorizontalAlign.Left ||
              style.verticalAlignment != VerticalAlign.Bottom ||
              style.wrap != null);

      if (hasAlignment) {
        buf.write(' applyAlignment="1"');
      }

      if (hasAlignment) {
        buf.write('><alignment');
        buf.write(
            ' horizontal="${style!.horizontalAlignment.toString().split('.').last.toLowerCase()}"');
        buf.write(
            ' vertical="${style.verticalAlignment.toString().split('.').last.toLowerCase()}"');
        if (style.wrap == TextWrapping.WrapText) {
          buf.write(' wrapText="1"');
        } else if (style.wrap == TextWrapping.Clip) {
          buf.write(' shrinkToFit="1"');
        }
        buf.write('/></xf>');
      } else {
        buf.write('/>');
      }
    }
    buf.write('</cellXfs>');

    buf.write('</styleSheet>');
    return buf.toString();
  }

  void _writeBorderSide(StringBuffer buf, String side, Border border) {
    final style = border.borderStyle;
    final color = border.borderColorHex;

    if (style == null && color == null) {
      buf.write('<$side/>');
    } else {
      buf.write('<$side');
      if (style != null) buf.write(' style="${style.style}"');
      buf.write('>');
      if (color != null) {
        buf.write('<color rgb="$color"/>');
      }
      buf.write('</$side>');
    }
  }

  String _buildSharedStrings() {
    final count = _sharedStringCount;
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="$count" uniqueCount="$count">');

    // Build sorted list by index
    final sorted = List<String?>.filled(count, null);
    _sharedStrings.forEach((text, idx) {
      sorted[idx] = text;
    });

    for (final text in sorted) {
      buf.write('<si><t xml:space="preserve">');
      buf.write(_escapeXmlStream(text ?? ''));
      buf.write('</t></si>');
    }

    buf.write('</sst>');
    return buf.toString();
  }

  String _buildWorksheet(_StreamSheet sheet) {
    final buf = StringBuffer();
    buf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    buf.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
    buf.write('<sheetData>');
    buf.write(sheet.buffer);
    buf.write('</sheetData>');
    buf.write('</worksheet>');
    return buf.toString();
  }

  static String _escapeXmlStream(String text) {
    if (!text.contains('&') &&
        !text.contains('<') &&
        !text.contains('>') &&
        !text.contains('"') &&
        !text.contains("'")) {
      return text;
    }
    return text
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&apos;');
  }
}

/// Internal per-sheet state for the streaming writer.
class _StreamSheet {
  final String name;
  final StringBuffer buffer = StringBuffer();
  int rowCount = 0;

  _StreamSheet(this.name);
}

/// Internal cellXf record for style deduplication.
class _StreamXf extends Equatable {
  final int fontId;
  final int fillId;
  final int borderId;
  final CellStyle? style;

  const _StreamXf({
    required this.fontId,
    required this.fillId,
    required this.borderId,
    required this.style,
  });

  @override
  List<Object?> get props => [fontId, fillId, borderId, style];
}

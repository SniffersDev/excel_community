part of '../../../excel_community.dart';

class _WorksheetManager {
  final Excel _excel;
  final Save _save;

  _WorksheetManager(this._excel, this._save);

  /// Sets up worksheet elements. Row/cell data is written directly to
  /// [sheetDataBuffers] as XML strings (keyed by the xmlSheetId) instead
  /// of building XmlElement DOM nodes. This dramatically reduces memory
  /// allocations for large files.
  void setSheetElements(
      {required Map<String, StringBuffer> sheetDataBuffers}) {
    _excel._sharedStrings.clear();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (_excel._sheets[sheetName] == null) {
        _save.parser._createSheet(sheetName);
      }

      // Clear existing DOM children of sheetData — we'll write directly
      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();
      }

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      double? defaultRowHeight = sheetObject.defaultRowHeight;
      double? defaultColumnWidth = sheetObject.defaultColumnWidth;

      XmlElement worksheetElement = xmlFile.findAllElements('worksheet').first;

      XmlElement? sheetFormatPrElement =
          worksheetElement.findElements('sheetFormatPr').isNotEmpty
              ? worksheetElement.findElements('sheetFormatPr').first
              : null;

      if (sheetFormatPrElement != null) {
        sheetFormatPrElement.attributes.clear();

        if (defaultRowHeight == null && defaultColumnWidth == null) {
          worksheetElement.children.remove(sheetFormatPrElement);
        }
      } else if (defaultRowHeight != null || defaultColumnWidth != null) {
        sheetFormatPrElement = XmlElement(XmlName('sheetFormatPr'), [], []);
        worksheetElement.children.insert(0, sheetFormatPrElement);
      }

      if (defaultRowHeight != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultRowHeight'), defaultRowHeight.toStringAsFixed(2)));
      }
      if (defaultColumnWidth != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultColWidth'), defaultColumnWidth.toStringAsFixed(2)));
      }

      _setColumns(sheetObject, xmlFile);

      // Write row/cell data to StringBuffer instead of DOM
      final xmlSheetId = _excel._xmlSheetId[sheetName];
      if (xmlSheetId != null) {
        final buffer = StringBuffer();
        _writeRowsToBuffer(sheetName, sheetObject, buffer);
        sheetDataBuffers[xmlSheetId] = buffer;
      }

      _save._setHeaderFooter(sheetName);
    });
  }

  void _setColumns(Sheet sheetObject, XmlDocument xmlFile) {
    final columnElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColumnWidths.isEmpty &&
        sheetObject.getColumnAutoFits.isEmpty) {
      if (columnElements.isEmpty) {
        return;
      }

      final columns = columnElements.first;
      final worksheet = xmlFile.findAllElements('worksheet').first;
      worksheet.children.remove(columns);
      return;
    }

    if (columnElements.isEmpty) {
      final worksheet = xmlFile.findAllElements('worksheet').first;
      final sheetData = xmlFile.findAllElements('sheetData').first;
      final index = worksheet.children.indexOf(sheetData);

      worksheet.children.insert(index, XmlElement(XmlName('cols'), [], []));
    }

    var columns = columnElements.first;

    if (columns.children.isNotEmpty) {
      columns.children.clear();
    }

    final autoFits = sheetObject.getColumnAutoFits;
    final customWidths = sheetObject.getColumnWidths;

    final columnCount = max(
        autoFits.isEmpty ? 0 : autoFits.keys.reduce(max) + 1,
        customWidths.isEmpty ? 0 : customWidths.keys.reduce(max) + 1);

    double defaultColumnWidth =
        sheetObject.defaultColumnWidth ?? _excelDefaultColumnWidth;

    for (var index = 0; index < columnCount; index++) {
      double width = defaultColumnWidth;

      if (autoFits.containsKey(index) && (!customWidths.containsKey(index))) {
        width = _calcAutoFitColumnWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          width = customWidths[index]!;
        }
      }

      _addNewColumn(columns, index, index, width);
    }
  }

  /// Writes row and cell XML data directly to a [StringBuffer] instead
  /// of building XmlElement DOM nodes. This is the key optimization for
  /// large files — it avoids creating millions of XmlElement objects.
  void _writeRowsToBuffer(
      String sheetName, Sheet sheetObject, StringBuffer buffer) {
    final customHeights = sheetObject.getRowHeights;

    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      if (sheetObject._sheetData[rowIndex] == null) {
        continue;
      }

      double? height;
      if (customHeights.containsKey(rowIndex)) {
        height = customHeights[rowIndex];
      }

      // Write <row> opening tag
      buffer.write('<row r="');
      buffer.write(rowIndex + 1);
      buffer.write('"');
      if (height != null) {
        buffer.write(' ht="');
        buffer.write(height.toStringAsFixed(2));
        buffer.write('" customHeight="1"');
      }
      buffer.write('>');

      for (var columnIndex = 0;
          columnIndex < sheetObject._maxColumns;
          columnIndex++) {
        var data = sheetObject._sheetData[rowIndex]![columnIndex];
        if (data == null) {
          continue;
        }
        _writeCellToBuffer(
            sheetName, buffer, columnIndex, rowIndex, data.value,
            data.cellStyle?.numberFormat);
      }

      buffer.write('</row>');
    }
  }

  /// Writes a single cell's XML directly to the [buffer].
  void _writeCellToBuffer(String sheet, StringBuffer buffer, int columnIndex,
      int rowIndex, CellValue? value, NumFormat? numberFormat) {
    int? sharedStringIndex;
    if (value is TextCellValue) {
      final existing = _excel._sharedStrings.tryFind(value.toString());
      if (existing != null) {
        sharedStringIndex = _excel._sharedStrings.add(existing, value.toString());
      } else {
        sharedStringIndex = _excel._sharedStrings.addFromString(value.toString());
      }
    }

    String rC = getCellId(columnIndex, rowIndex);

    // Write <c> opening tag with attributes
    buffer.write('<c r="');
    buffer.write(rC);
    buffer.write('"');

    // Style attribute
    final cellStyle =
        _excel._sheetMap[sheet]?._sheetData[rowIndex]?[columnIndex]?.cellStyle;

    int? styleIndex;
    if (_excel._styleChanges && cellStyle != null) {
      int upperLevelPos = _checkPosition(_excel._cellStyleList, cellStyle);
      if (upperLevelPos == -1) {
        int lowerLevelPos = _checkPosition(_save._innerCellStyle, cellStyle);
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      styleIndex = upperLevelPos;
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      styleIndex = _excel._cellStyleReferenced[sheet]![rC];
    }

    if (styleIndex != null) {
      buffer.write(' s="');
      buffer.write(styleIndex);
      buffer.write('"');
    }

    if (value is TextCellValue) {
      buffer.write(' t="s"');
    } else if (value is BoolCellValue) {
      buffer.write(' t="b"');
    }

    buffer.write('>');

    // Write cell value children
    switch (value) {
      case null:
        break;
      case TextCellValue():
        buffer.write('<v>');
        buffer.write(sharedStringIndex);
        buffer.write('</v>');
        break;
      case FormulaCellValue():
        buffer.write('<f>');
        buffer.write(_escapeXml(value.formula));
        buffer.write('</f><v>');
        buffer.write(_escapeXml(value.write(numberFormat)));
        buffer.write('</v>');
        break;
      case IntCellValue() ||
            DoubleCellValue() ||
            DateCellValue() ||
            TimeCellValue() ||
            DateTimeCellValue() ||
            BoolCellValue():
        buffer.write('<v>');
        buffer.write(_escapeXml(value.write(numberFormat)));
        buffer.write('</v>');
        break;
    }

    buffer.write('</c>');
  }

  /// Escape special XML characters in text content.
  static String _escapeXml(String text) {
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

  void _addNewColumn(XmlElement columns, int min, int max, double width) {
    columns.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
      XmlAttribute(XmlName('bestFit'), "1"),
      XmlAttribute(XmlName('customWidth'), "1"),
    ], []));
  }

  double _calcAutoFitColumnWidth(Sheet sheet, int column) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(column) &&
          value[column]!.value is! FormulaCellValue) {
        maxNumOfCharacters =
            max(value[column]!.value.toString().length, maxNumOfCharacters);
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }
}

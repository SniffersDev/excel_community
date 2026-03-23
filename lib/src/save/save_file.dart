part of excel_community;

class Save {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};
  final List<CellStyle> _innerCellStyle = [];
  final Parser parser;
  late final _ChartManager _chartManager;
  late final _StyleManager _styleManager;
  late final _WorksheetManager _worksheetManager;
  late final _WorkbookManager _workbookManager;

  Save._(this._excel, this.parser) {
    _chartManager = _ChartManager(_excel, this);
    _styleManager = _StyleManager(_excel, this);
    _worksheetManager = _WorksheetManager(_excel, this);
    _workbookManager = _WorkbookManager(_excel, this);
  }

  List<int>? _save() {
    if (_excel._styleChanges) {
      _styleManager.processStylesFile();
    }

    // Collect sheet data that will be written directly to StringBuffer
    // instead of building XmlElement DOM nodes.
    final sheetDataBuffers = <String, StringBuffer>{};
    _worksheetManager.setSheetElements(sheetDataBuffers: sheetDataBuffers);

    if (_excel._defaultSheet != null) {
      _workbookManager.setDefaultSheet(_excel._defaultSheet);
    }

    _workbookManager.setSharedStrings();

    if (_excel._mergeChanges) {
      _workbookManager.setMerge();
    }

    if (_excel._rtlChanges) {
      _workbookManager.setRTL();
    }

    _chartManager.processCharts();

    // Collect the set of worksheet XML file keys for streaming serialization
    final worksheetXmlIds =
        Set<String>.from(_excel._xmlSheetId.values);

    for (var xmlFile in _excel._xmlFiles.keys) {
      List<int> content;
      if (worksheetXmlIds.contains(xmlFile) &&
          sheetDataBuffers.containsKey(xmlFile)) {
        // For worksheet files, serialize around <sheetData> using the
        // pre-built StringBuffer to avoid huge DOM .toString() calls.
        content = _serializeWorksheetXml(
            _excel._xmlFiles[xmlFile]!, sheetDataBuffers[xmlFile]!);
      } else {
        content = utf8.encode(_excel._xmlFiles[xmlFile].toString());
      }
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_excel._archive, _archiveFiles));
  }

  /// Serialize a worksheet XML document by splitting around <sheetData>.
  /// The [sheetDataBuffer] contains the pre-built row/cell XML written
  /// directly by the WorksheetManager, avoiding millions of XmlElement
  /// allocations for large sheets.
  List<int> _serializeWorksheetXml(
      XmlDocument xmlDoc, StringBuffer sheetDataBuffer) {
    final fullXml = xmlDoc.toString();

    // Find the <sheetData.../> or <sheetData...>...</sheetData> region
    final sheetDataOpenIdx = fullXml.indexOf('<sheetData');
    if (sheetDataOpenIdx == -1) {
      // No sheetData found — fallback to normal serialization
      return utf8.encode(fullXml);
    }

    // Find end of the opening tag
    final afterOpen = fullXml.indexOf('>', sheetDataOpenIdx);
    if (afterOpen == -1) {
      return utf8.encode(fullXml);
    }

    // Check if self-closing: <sheetData/>
    final isSelfClosing = fullXml[afterOpen - 1] == '/';

    final buf = StringBuffer();
    if (isSelfClosing) {
      // Replace <sheetData/> with <sheetData>...rows...</sheetData>
      buf.write(fullXml.substring(0, sheetDataOpenIdx));
      buf.write('<sheetData>');
      buf.write(sheetDataBuffer);
      buf.write('</sheetData>');
      buf.write(fullXml.substring(afterOpen + 1));
    } else {
      // Replace content between <sheetData> and </sheetData>
      final closeIdx = fullXml.indexOf('</sheetData>', afterOpen);
      if (closeIdx == -1) {
        return utf8.encode(fullXml);
      }
      buf.write(fullXml.substring(0, afterOpen + 1));
      buf.write(sheetDataBuffer);
      buf.write(fullXml.substring(closeIdx));
    }

    return utf8.encode(buf.toString());
  }

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
        leftBorder: cellStyle.leftBorder,
        rightBorder: cellStyle.rightBorder,
        topBorder: cellStyle.topBorder,
        bottomBorder: cellStyle.bottomBorder,
        diagonalBorder: cellStyle.diagonalBorder,
        diagonalBorderUp: cellStyle.diagonalBorderUp,
        diagonalBorderDown: cellStyle.diagonalBorderDown,
      );

  void _setHeaderFooter(String sheetName) {
    _workbookManager.setHeaderFooter(sheetName);
  }

  void _addContentType(String contentType, String partName) {
    final contentTypes = _excel._xmlFiles['[Content_Types].xml'];
    if (contentTypes == null) return;

    final typesElement = contentTypes.findAllElements('Types').first;

    // Check if already exists
    final exists = typesElement.children.any((node) =>
        node is XmlElement && node.getAttribute('PartName') == partName);

    if (!exists) {
      typesElement.children.add(XmlElement(XmlName('Override'), [
        XmlAttribute(XmlName('PartName'), partName),
        XmlAttribute(XmlName('ContentType'), contentType),
      ]));
    }
  }
}

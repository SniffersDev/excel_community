part of excel_community;

/// A single cell in a streaming Excel row.
class ExcelCell {
  /// Zero-based column index.
  final int columnIndex;

  /// The cell value.
  final CellValue value;

  const ExcelCell(this.columnIndex, this.value);

  @override
  String toString() => 'ExcelCell($columnIndex, $value)';
}

/// A single row yielded by [ExcelStreamReader].
class ExcelRow {
  /// Zero-based row index.
  final int index;

  /// The cells in this row (may be sparse — only non-empty cells).
  final List<ExcelCell> cells;

  const ExcelRow(this.index, this.cells);

  /// Get cell value by column index, or null if not present.
  CellValue? cellAt(int columnIndex) {
    for (final cell in cells) {
      if (cell.columnIndex == columnIndex) return cell.value;
    }
    return null;
  }

  @override
  String toString() => 'ExcelRow($index, ${cells.length} cells)';
}

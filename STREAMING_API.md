# Streaming API Guide

The Streaming API provides memory-efficient Excel reading and writing for large files (400k+ rows).  
Unlike the DOM-based `Excel` class, it never builds the full spreadsheet in memory, making it ideal for **browser environments** and **resource-constrained** targets.

> [!NOTE]
> The Streaming API supports **text**, **int**, **double**, **bool**, **date**, **time**, and **datetime** cell types.  
> It does **not** support charts, merged cells, styles, or formulas.

---

## Writing — `ExcelStreamWriter`

### Quick start

```dart
import 'package:excel_community/excel_community.dart';

final writer = ExcelStreamWriter(sheetName: 'Sales');

// Header row (convenience — wraps every value in TextCellValue)
writer.addHeaderRow(['Month', 'Revenue', 'Profit']);

// Data rows
writer.addRow([TextCellValue('Jan'), IntCellValue(50000), DoubleCellValue(12500.75)]);
writer.addRow([TextCellValue('Feb'), IntCellValue(62000), DoubleCellValue(15800.00)]);
writer.addRow([TextCellValue('Mar'), IntCellValue(48000), DoubleCellValue(11200.50)]);

// Encode to XLSX bytes
final bytes = writer.encode();
```

### Multiple sheets

```dart
final writer = ExcelStreamWriter(sheetName: 'Employees');

writer.addHeaderRow(['Name', 'Department', 'Active']);
writer.addRow([TextCellValue('Alice'), TextCellValue('Engineering'), BoolCellValue(true)]);
writer.addRow([TextCellValue('Bob'),   TextCellValue('Marketing'),   BoolCellValue(true)]);

// Switch to a new sheet — subsequent addRow() calls target it
writer.addSheet('Departments');
writer.addHeaderRow(['Department', 'Headcount']);
writer.addRow([TextCellValue('Engineering'), IntCellValue(42)]);
writer.addRow([TextCellValue('Marketing'),   IntCellValue(18)]);

final bytes = writer.encode();
```

### Browser download (web only)

```dart
// Triggers a browser file-download dialog
writer.save(fileName: 'report.xlsx');
```

### Generating large datasets

```dart
final writer = ExcelStreamWriter(sheetName: 'LargeData');
writer.addHeaderRow(['Row', 'Value', 'Score']);

for (var i = 0; i < 400000; i++) {
  writer.addRow([
    IntCellValue(i),
    TextCellValue('item_$i'),
    DoubleCellValue(i * 0.01),
  ]);
}

final bytes = writer.encode();
print('Generated ${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB');
```

> [!IMPORTANT]
> `encode()` can only be called **once**. After encoding, no more rows or sheets can be added.

---

## Reading — `ExcelStreamReader`

### Quick start

```dart
import 'package:excel_community/excel_community.dart';

final reader = ExcelStreamReader.fromBytes(xlsxBytes);

// List available sheets
print(reader.sheetNames); // e.g. ['Sales', 'Departments']

// Iterate rows lazily — only one row is in memory at a time
for (final row in reader.readSheet('Sales')) {
  print('Row ${row.index}: ${row.cells.map((c) => c.value).toList()}');
}
```

### Accessing cell values

Each `ExcelRow` contains a list of `ExcelCell` objects with a `columnIndex` and a `value`:

```dart
for (final row in reader.readSheet('Sheet1')) {
  // By iterating cells
  for (final cell in row.cells) {
    print('  Column ${cell.columnIndex}: ${cell.value}');
  }

  // By column index (returns null if cell is empty)
  final name  = row.cellAt(0); // CellValue?
  final score = row.cellAt(2); // CellValue?
}
```

### Type-safe value extraction

```dart
for (final row in reader.readSheet('Data')) {
  final value = row.cellAt(0);

  if (value is TextCellValue) {
    print('Text: $value');
  } else if (value is IntCellValue) {
    print('Int: $value');
  } else if (value is DoubleCellValue) {
    print('Double: $value');
  } else if (value is BoolCellValue) {
    print('Bool: $value');
  }
}
```

### Reading multiple sheets

```dart
final reader = ExcelStreamReader.fromBytes(xlsxBytes);

for (final sheetName in reader.sheetNames) {
  print('--- $sheetName ---');
  var rowCount = 0;
  for (final row in reader.readSheet(sheetName)) {
    rowCount++;
  }
  print('$rowCount rows');
}
```

### Interoperability with the DOM-based `Excel` class

Files created with `ExcelStreamWriter` can be read by `Excel.decodeBytes()`, and files created with `Excel` can be read by `ExcelStreamReader`:

```dart
// Write with Stream → Read with DOM
final writer = ExcelStreamWriter(sheetName: 'Data');
writer.addRow([TextCellValue('hello'), IntCellValue(42)]);
final bytes = writer.encode();

final excel = Excel.decodeBytes(bytes);
print(excel.tables['Data']!.rows[0][0]!.value); // hello

// Write with DOM → Read with Stream
final domExcel = Excel.createExcel();
domExcel['Sheet1'].updateCell(CellIndex.indexByString('A1'), TextCellValue('world'));
final domBytes = domExcel.encode()!;

final reader = ExcelStreamReader.fromBytes(domBytes);
final rows = reader.readSheet('Sheet1').toList();
print(rows[0].cells[0].value); // world
```

---

## API Reference

### `ExcelStreamWriter`

| Method / Constructor | Description |
|---|---|
| `ExcelStreamWriter({String sheetName})` | Create a writer with an initial sheet (default `'Sheet1'`). |
| `addSheet(String name)` | Add a new sheet and make it the active target. Returns `this` for chaining. |
| `addHeaderRow(List<String> headers)` | Convenience to add a row of `TextCellValue` headers. |
| `addRow(List<CellValue?> values)` | Append a row of typed cell values to the active sheet. |
| `encode()` | Produce the final XLSX bytes. **One-shot** — cannot be called twice. |
| `save({String fileName})` | Web only — calls `encode()` then triggers a browser download. |

### `ExcelStreamReader`

| Method / Constructor | Description |
|---|---|
| `ExcelStreamReader.fromBytes(List<int> bytes)` | Create a reader from XLSX bytes. |
| `sheetNames` | `List<String>` of available sheet names. |
| `readSheet(String name)` | Returns a lazy `Iterable<ExcelRow>` — rows are parsed on demand. |

### `ExcelRow`

| Member | Description |
|---|---|
| `index` | Zero-based row index. |
| `cells` | `List<ExcelCell>` — only non-empty cells (may be sparse). |
| `cellAt(int columnIndex)` | Returns the `CellValue?` at the given column, or `null`. |

### `ExcelCell`

| Member | Description |
|---|---|
| `columnIndex` | Zero-based column index. |
| `value` | The `CellValue` of the cell. |

---

## When to use Streaming vs DOM

| | `Excel` (DOM) | `ExcelStreamWriter` / `ExcelStreamReader` |
|---|---|---|
| **Best for** | Small–medium files, full editing | Large files, append-only writes, linear reads |
| **Memory** | Entire workbook held in memory | Row-at-a-time |
| **Features** | Styles, formulas, merges, charts | Cell values only |
| **Row limit** | Bounded by available RAM | 400k+ rows tested |
| **Interop** | ↔ Stream API (read/write) | ↔ DOM API (read/write) |

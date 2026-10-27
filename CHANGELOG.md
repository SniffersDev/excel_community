# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2026-10-27
### Added
- `ZipStreamWriter`: writes a zip to a `ByteSink` entry by entry while the content is produced (data descriptors, stored or deflated entries, zip64 when an entry or the archive passes 4 GiB, `abandonEntry` for a content stream that fails halfway).
- `ExcelStreamWriter.streamed(zip)`: writes each sheet into a `ZipStreamWriter` as rows arrive (every `flushRows` rows or 1 MiB), text cells inline (`inlineStrings`, default on) so no string table grows with the rows, and `finish()` writes the workbook parts last. A workbook of any size costs the memory of its pending rows.
- `StreamDeflate` (internal): the archive package's deflate with chunked input (`addBytes` keeps the stream open, `finish()` closes it); the package's own `Deflate.addBytes` ends the stream on every call.

## [1.1.1] - 2026-10-27
### Changed
- `ExcelStreamReader` loads the shared strings by scanning `<si>` items on the bytes instead of parsing the whole `sharedStrings.xml` into a DOM: a 10 MB table no longer costs a DOM of the whole part, and the time before the first row drops with it.
- `readSheetAsync` awaits `pause` after inflating the sheet and after loading the shared strings, so a UI thread gets a frame before the first row.

## [1.1.0] - 2026-10-27
### Added
- `ExcelStreamReader.readSheetAsync(sheet, {pause, sliceSize})`: the same rows as `readSheet` as a `Stream`, awaiting `pause` after every `sliceSize` bytes (64 KiB) so a UI thread can repaint while a large sheet is read.
- `ExcelStreamReader.rowCount(sheet)`: the row count from the sheet's `<dimension>`, for progress; null when absent.
- `ExcelStreamReader` reads dates: a numeric cell whose style is a date or time number format (built-in ids 14–22, 45–47 and custom date codes) becomes `DateCellValue`, `DateTimeCellValue` or `TimeCellValue`, as `Excel.decodeBytes` does. Before, every number came back as `IntCellValue` / `DoubleCellValue`.
- `ExcelStreamReader` reads inline strings (`t="inlineStr"`, rich runs concatenated) and formula string results (`t="str"`); error cells are skipped.
### Changed
- `ExcelStreamReader` scans a sheet on its decompressed bytes, one `<row>` at a time, instead of decoding the whole sheet to a string and taking a `substring` of `<sheetData>`. Memory is bounded by the sheet's bytes plus one row: the 257 MB Measurements sheet of a 39 MB workbook (279,526 rows) reads in 15 s inside a 1 GB heap, where the previous version needed two copies of the sheet text.

## [1.0.7] - 2026-10-20
- Cell styling support in `ExcelStreamWriter` (fonts, fills, borders, alignment).

## [1.0.6] - 2026-03-17
- Fix color accessibility.
## [1.0.5] - 2026-03-17
- Apply clean code fix minors bugs.
## [1.0.4] - 2026-03-15
- Apply Clean Architecture in Save class.
- Refactor Sheet class.
- Refactor Color class.
- Refactor Border class.
- Refactor Formula class.
- Refactor NumberFormat class.

## [1.0.3] - 2026-03-15

- Renamed main example to `example/main.dart` for better visibility on pub.dev.
- Significantly increased documentation coverage (>28%) across the public API.
- Added detailed documentation for all chart types and cell value classes.

## [1.0.2] - 2026-03-15

## [1.0.1] - 2026-03-14

- Reorganized tests: moved manual debug scripts to `test/manual/`.
- Fixed static analysis errors in `parse.dart` and `save_file.dart`.
- Improved package metadata for pub.dev.

## [1.0.0] - 2026-03-14

- Initial Release of the community fork.
- Forked from [excel](https://github.com/justkawal/excel)
- Updated minimum Dart SDK to 3.6.0
- Migrate from `dart:html` to `package:web`
- Update packages to latest versions
- Added support for charts


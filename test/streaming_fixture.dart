import 'dart:typed_data';

import 'package:archive/archive.dart';

/// Builds a minimal .xlsx from hand-written sheet XML, so reader tests do not
/// depend on any writer. [sheetXml] maps a sheet name to the content of its
/// `<sheetData>` (the `<row>` elements); [dimension] adds a `<dimension>`
/// element per sheet when given.
Uint8List buildXlsx({
  required Map<String, String> sheetXml,
  String? sharedStringsXml,
  String? stylesXml,
  Map<String, String> dimension = const {},
  bool storeSheets = false,
}) {
  final archive = Archive();
  void add(String path, String content, {bool store = false}) => archive.addFile(
      ArchiveFile.string(path, content)
        ..compression = store ? CompressionType.none : CompressionType.deflate);

  final names = sheetXml.keys.toList();
  add(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '${names.asMap().entries.map((e) => '<Override PartName="/xl/worksheets/sheet${e.key + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join()}'
    '</Types>',
  );
  add(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>',
  );
  add(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets>'
    '${names.asMap().entries.map((e) => '<sheet name="${e.value}" sheetId="${e.key + 1}" r:id="rId${e.key + 1}"/>').join()}'
    '</sheets></workbook>',
  );
  add(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '${names.asMap().entries.map((e) => '<Relationship Id="rId${e.key + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${e.key + 1}.xml"/>').join()}'
    '</Relationships>',
  );
  for (final entry in names.asMap().entries) {
    final dim = dimension[entry.value];
    add(
      'xl/worksheets/sheet${entry.key + 1}.xml',
      store: storeSheets,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      '${dim == null ? '' : '<dimension ref="$dim"/>'}'
      '<sheetData>${sheetXml[entry.value]}</sheetData>'
      '</worksheet>',
    );
  }
  if (sharedStringsXml != null) {
    add(
      'xl/sharedStrings.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">$sharedStringsXml</sst>',
    );
  }
  if (stylesXml != null) {
    add(
      'xl/styles.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">$stylesXml</styleSheet>',
    );
  }
  return Uint8List.fromList(ZipEncoder().encode(archive));
}

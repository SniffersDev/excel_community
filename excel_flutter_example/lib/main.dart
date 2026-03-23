import 'dart:io';
import 'dart:js_interop';
import 'dart:typed_data';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart' hide Border, BorderStyle;
import 'package:excel_community/excel_community.dart';
import 'package:file_picker/file_picker.dart';
import 'package:web/web.dart' as web;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Chart Example',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'Excel Chart Example'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String _status = 'Press a button to generate an Excel with a chart.';
  bool _isGenerating = false;

  Future<void> _generateSimpleExcel() async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating simple Excel (NO chart)...';
    });

    try {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Add simple data
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Age"));
      sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alice"));
      sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(30));
      sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Bob"));
      sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(25));

      if (kIsWeb) {
        if (kDebugMode) {
          print('Generating simple Excel for Web...');
        }
        
        final bytes = excel.save(fileName: 'simple_no_chart.xlsx');
        
        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Simple Excel generated successfully!\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
                'The download should start automatically.\n'
                '\n📥 Check your Downloads folder\n'
                '📌 File: simple_no_chart.xlsx';
          });
          
          if (kDebugMode) {
            print('✅ Simple Excel saved for web: ${bytes.length} bytes');
          }
        } else {
          setState(() {
            _status = '❌ Error: Failed to generate Excel file.\n'
                'The file is empty or could not be encoded.';
          });
        }
      } else {
        var bytes = excel.encode();
        
        if (bytes == null) {
          setState(() {
            _status = 'Error: Failed to encode Excel file.';
          });
          return;
        }
        
        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Simple Excel File',
          fileName: 'simple_no_chart.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
          await file.writeAsBytes(bytes);
          
          final savedFileSize = await file.length();
          
          setState(() {
            _status = '✅ Simple Excel saved successfully!\n'
                'Location: $outputFile\n'
                'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
                '\n📌 Open with Excel to verify it works';
          });
        } else {
          setState(() {
            _status = 'Save cancelled.';
          });
        }
      }
    } catch (e, stackTrace) {
      setState(() {
        _status = 'Error: $e';
      });
      if (kDebugMode) {
        print('Error: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() {
        _isGenerating = false;
      });
    }
  }

  Future<void> _generateExcelWithChart(ChartType type) async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Excel with chart...';
    });

    try {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Add Headers
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value 1"));
      sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Value 2"));

      // Add Data
      final data = [
        ['Jan', 10, 15],
        ['Feb', 20, 25],
        ['Mar', 15, 30],
        ['Apr', 25, 20],
        ['May', 30, 35],
        ['Jun', 20, 40],
      ];

      for (var i = 0; i < data.length; i++) {
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(data[i][0] as String));
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(data[i][1] as int));
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(data[i][2] as int));
      }

      Chart chart;
      final series = [
        ChartSeries(
          name: "Series 1",
          categoriesRange: r"Sheet1!$A$2:$A$7",
          valuesRange: r"Sheet1!$B$2:$B$7",
        ),
        ChartSeries(
          name: "Series 2",
          categoriesRange: r"Sheet1!$A$2:$A$7",
          valuesRange: r"Sheet1!$C$2:$C$7",
        ),
      ];

      final anchor = ChartAnchor.at(column: 5, row: 1, width: 10, height: 15);

      switch (type) {
        case ChartType.column:
          chart = ColumnChart(
            title: "Monthly Data (Column)",
            series: series,
            anchor: anchor,
          );
        case ChartType.bar:
          chart = BarChart(
            title: "Monthly Data (Bar)",
            series: series,
            anchor: anchor,
          );
        case ChartType.line:
          chart = LineChart(
            title: "Monthly Data (Line)",
            series: series,
            anchor: anchor,
          );
        case ChartType.area:
          chart = AreaChart(
            title: "Monthly Data (Area)",
            series: series,
            anchor: anchor,
          );
        case ChartType.pie:
          chart = PieChart(
            title: "Pie Chart Example",
            series: [series[0]], // Pie chart usually only takes one series
            anchor: anchor,
          );
        case ChartType.doughnut:
          chart = DoughnutChart(
            title: "Doughnut Chart Example",
            series: [series[0]], // Doughnut chart usually only takes one series
            anchor: anchor,
          );
        case ChartType.radar:
          chart = RadarChart(
            title: "Radar Chart Example",
            series: series,
            anchor: anchor,
            filled: true,
          );
        case ChartType.scatter:
          chart = ScatterChart(
            title: "Scatter Chart Example",
            series: series,
            anchor: anchor,
          );
      }

      sheet.addChart(chart);
      
      if (kDebugMode) {
        print('Chart added to sheet');
      }

      if (kIsWeb) {
        if (kDebugMode) {
          print('Generating Excel for Web...');
        }
        
        final bytes = excel.save(fileName: 'chart_example.xlsx');
        
        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Excel generated successfully!\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
                'The download should start automatically.\n'
                '\n📥 Check your Downloads folder\n'
                '📌 File: chart_example.xlsx';
          });
          
          if (kDebugMode) {
            print('✅ Excel saved for web: ${bytes.length} bytes');
          }
        } else {
          setState(() {
            _status = '❌ Error: Failed to generate Excel file for web.\n'
                'The file is empty or could not be encoded.';
          });
          
          if (kDebugMode) {
            print('❌ Error: excel.save() returned null or empty');
          }
        }
      } else {
        var bytes = excel.encode();
        
        if (bytes == null) {
          setState(() {
            _status = 'Error: Failed to encode Excel file.';
          });
          return;
        }
        
        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Excel File',
          fileName: 'chart_example.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
        
          await file.writeAsBytes(bytes);
          
          // Verify the file was written correctly
          final savedFileSize = await file.length();
          
          setState(() {
            _status = '✅ Excel saved successfully!\n'
                'Location: $outputFile\n'
                'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
                '\n📌 IMPORTANT: Open with Excel, not a text editor!';
          });
          
          if (kDebugMode) {
            print('✅ File saved: $outputFile');
            print('📊 Bytes generated: ${bytes.length}');
            print('💾 File size: $savedFileSize bytes');
          }
        } else {
          setState(() {
            _status = 'Save cancelled.';
          });
        }
      }
    } catch (e, stackTrace) {
      setState(() {
        _status = 'Error: $e';
      });
      if (kDebugMode) {
        print('Error generating Excel with chart: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() {
        _isGenerating = false;
      });
    }
  }

  Future<void> _generateFullExample() async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Full Demo Excel...';
    });

    try {
      var excel = Excel.createExcel();
      // Rename default sheet
      var sheetName = 'Full Demo';
      excel.rename('Sheet1', sheetName);
      var sheet = excel[sheetName];

      // 1. Defining Styles with Borders and Colors
      final headerStyle = CellStyle(
        bold: true,
        fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
        backgroundColorHex: ExcelColor.fromHexString('#4472C4'),
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
        rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
        topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
        bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
      );

      final dataStyle = CellStyle(
        leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
        rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
        topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
        bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
      );

      final formulaStyle = CellStyle(
        bold: true,
        backgroundColorHex: ExcelColor.fromHexString('#E2EFDA'),
        leftBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
        rightBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
        topBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
        bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
      );

      final multiColorStyle = CellStyle(
        bold: true,
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        leftBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#00FF00')), // Green
        rightBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#FFFF00')), // Yellow
        topBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#FF0000')), // Red
        bottomBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#0000FF')), // Blue
      );

      // 2. Headers
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"), cellStyle: headerStyle);
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Revenue"), cellStyle: headerStyle);
      sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Expenses"), cellStyle: headerStyle);
      sheet.updateCell(CellIndex.indexByString("D1"), TextCellValue("Profit"), cellStyle: headerStyle);

      // 3. Data and Formulas
      final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
      final revenues = [1000, 1200, 1500, 1300, 1700, 2000];
      final expenses = [800, 900, 1000, 950, 1100, 1200];

      for (var i = 0; i < months.length; i++) {
        var row = i + 2;
        sheet.updateCell(CellIndex.indexByString("A$row"), TextCellValue(months[i]), cellStyle: dataStyle);
        sheet.updateCell(CellIndex.indexByString("B$row"), IntCellValue(revenues[i]), cellStyle: dataStyle);
        sheet.updateCell(CellIndex.indexByString("C$row"), IntCellValue(expenses[i]), cellStyle: dataStyle);
        // Formula: Revenue - Expenses
        sheet.updateCell(CellIndex.indexByString("D$row"), FormulaCellValue("B$row-C$row"), cellStyle: dataStyle);
      }

      // 4. Totals with Formulas
      var totalRow = months.length + 2;
      sheet.updateCell(CellIndex.indexByString("A$totalRow"), TextCellValue("TOTAL"), cellStyle: formulaStyle);
      sheet.updateCell(CellIndex.indexByString("B$totalRow"), FormulaCellValue("SUM(B2:B${totalRow - 1})"), cellStyle: formulaStyle);
      sheet.updateCell(CellIndex.indexByString("C$totalRow"), FormulaCellValue("SUM(C2:C${totalRow - 1})"), cellStyle: formulaStyle);
      sheet.updateCell(CellIndex.indexByString("D$totalRow"), FormulaCellValue("SUM(D2:D${totalRow - 1})"), cellStyle: formulaStyle);

      // 5. Multi-colored Border Demo
      sheet.updateCell(CellIndex.indexByString("A${totalRow + 2}"), 
          TextCellValue("Multi-colored Borders"), 
          cellStyle: multiColorStyle);
      sheet.setColumnWidth(0, 25.0);

      // 5. Chart using the data
      final series = [
        ChartSeries(
          name: "Revenue",
          categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
          valuesRange: "'Full Demo'!\$B\$2:\$B\$7",
        ),
        ChartSeries(
          name: "Profit",
          categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
          valuesRange: "'Full Demo'!\$D\$2:\$D\$7",
        ),
      ];

      final chart = ColumnChart(
        title: "Financial Overview",
        series: series,
        anchor: ChartAnchor.at(column: 6, row: 1, width: 10, height: 15),
      );

      sheet.addChart(chart);

      // 6. Save and Download
      if (kIsWeb) {
        final bytes = excel.save(fileName: 'full_demo_example.xlsx');
        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Full Demo Excel generated successfully!\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
                'The download should start automatically.\n'
                '📌 Includes: Formulas, Headers, Borders, and Charts.';
          });
        }
      } else {
        var bytes = excel.encode();
        if (bytes == null) {
          setState(() => _status = 'Error: Failed to encode Excel file.');
          return;
        }

        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Full Demo Excel',
          fileName: 'full_demo_example.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
          await file.writeAsBytes(bytes);
          setState(() {
            _status = '✅ Full Demo Excel saved successfully!\n'
                'Location: $outputFile\n'
                '📌 Includes: Formulas, Headers, Borders, and Charts.';
          });
        }
      }
    } catch (e, stackTrace) {
      setState(() => _status = 'Error: $e');
      if (kDebugMode) {
        print('Error generating full demo: $e');
        print(stackTrace);
      }
    } finally {
      setState(() => _isGenerating = false);
    }
  }

  Future<void> _generateLargeExcel(int rowCount) async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating large Excel ($rowCount rows × 20 cols)...';
    });

    try {
      final stopwatch = Stopwatch()..start();

      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      const cols = 20;

      for (var c = 0; c < cols; c++) {
        sheet
            .cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: c))
            .value = TextCellValue('Column_$c');
      }

      for (var r = 1; r <= rowCount; r++) {
        for (var c = 0; c < cols; c++) {
          final cell = sheet
              .cell(CellIndex.indexByColumnRow(rowIndex: r, columnIndex: c));
          switch (c % 4) {
            case 0:
              cell.value = TextCellValue('R${r}C$c');
              break;
            case 1:
              cell.value = IntCellValue(r * cols + c);
              break;
            case 2:
              cell.value = DoubleCellValue(r * 1.5 + c * 0.1);
              break;
            case 3:
              cell.value = IntCellValue(r + c);
              break;
          }
        }
      }

      final generateTime = stopwatch.elapsed;
      stopwatch.reset();

      if (kIsWeb) {
        final bytes = excel.save(fileName: 'large_${rowCount}_rows.xlsx');
        final encodeTime = stopwatch.elapsed;

        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Large Excel ($rowCount rows) generated!\n'
                'Generation: ${generateTime.inMilliseconds}ms\n'
                'Encode+Save: ${encodeTime.inMilliseconds}ms\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(0)} KB\n'
                '📥 Check Downloads folder';
          });
        } else {
          setState(() {
            _status = '❌ Failed: encode returned null/empty';
          });
        }
      } else {
        var bytes = excel.encode();
        final encodeTime = stopwatch.elapsed;

        if (bytes == null) {
          setState(() => _status = 'Error: Failed to encode.');
          return;
        }

        setState(() {
          _status = '✅ Large Excel ($rowCount rows) encoded!\n'
              'Generation: ${generateTime.inMilliseconds}ms\n'
              'Encode: ${encodeTime.inMilliseconds}ms\n'
              'File size: ${(bytes.length / 1024).toStringAsFixed(0)} KB';
        });
      }
    } catch (e, stackTrace) {
      setState(() {
        _status = '❌ CRASH with $rowCount rows!\n'
            'Error: $e\n'
            'Stack: ${stackTrace.toString().length > 300 ? stackTrace.toString().substring(0, 300) : stackTrace}';
      });
      if (kDebugMode) {
        print('Error generating large Excel: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() => _isGenerating = false);
    }
  }

  Future<void> _streamLargeExcel(int rowCount) async {
    setState(() {
      _isGenerating = true;
      _status = 'Stream writing $rowCount rows × 20 cols...\n'
          '⏳ Starting...';
    });

    // Yield to let the UI render the initial status
    await Future.delayed(Duration.zero);

    try {
      final stopwatch = Stopwatch()..start();
      const cols = 20;
      const batchSize = 10000; // yield to UI every 10k rows

      final writer = ExcelStreamWriter(sheetName: 'Sheet1');
      writer.addHeaderRow(List.generate(cols, (c) => 'Column_$c'));

      for (var r = 1; r <= rowCount; r++) {
        writer.addRow(List.generate(cols, (c) {
          switch (c % 4) {
            case 0:
              return TextCellValue('R${r}C$c');
            case 1:
              return IntCellValue(r * cols + c);
            case 2:
              return DoubleCellValue(r * 1.5 + c * 0.1);
            default:
              return IntCellValue(r + c);
          }
        }));

        // Yield to the UI every batchSize rows
        if (r % batchSize == 0) {
          final pct = (r * 100 / rowCount).toStringAsFixed(0);
          final elapsed = stopwatch.elapsed;
          setState(() {
            _status = 'Stream writing $rowCount rows × 20 cols...\n'
                '📝 $r / $rowCount rows ($pct%)\n'
                '⏱ Elapsed: ${elapsed.inSeconds}s';
          });
          await Future.delayed(Duration.zero);
        }
      }

      final generateTime = stopwatch.elapsed;
      setState(() {
        _status = 'Rows generated in ${generateTime.inSeconds}s\n'
            '📦 Encoding to XLSX...';
      });
      await Future.delayed(Duration.zero);
      stopwatch.reset();

      if (kIsWeb) {
        final bytes = writer.encode();
        final encodeTime = stopwatch.elapsed;

        // Trigger download via package:web (works in both JS and WASM)
        final uint8 = Uint8List.fromList(bytes);
        final blob = web.Blob(
          [uint8.toJS].toJS,
          web.BlobPropertyBag(
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          ),
        );
        final url = web.URL.createObjectURL(blob);
        final anchor = web.HTMLAnchorElement()
          ..href = url
          ..download = 'stream_${rowCount}_rows.xlsx';
        web.document.body?.append(anchor);
        anchor.click();
        anchor.remove();
        web.URL.revokeObjectURL(url);

        setState(() {
          _status = '✅ Stream Excel ($rowCount rows) generated!\n'
              'Generation: ${generateTime.inMilliseconds}ms\n'
              'Encode+Save: ${encodeTime.inMilliseconds}ms\n'
              'File size: ${(bytes.length / 1024).toStringAsFixed(0)} KB\n'
              '📥 Check Downloads folder';
        });
      } else {
        final bytes = writer.encode();
        final encodeTime = stopwatch.elapsed;

        setState(() {
          _status = '✅ Stream Excel ($rowCount rows) encoded!\n'
              'Generation: ${generateTime.inMilliseconds}ms\n'
              'Encode: ${encodeTime.inMilliseconds}ms\n'
              'File size: ${(bytes.length / 1024).toStringAsFixed(0)} KB';
        });
      }
    } catch (e, stackTrace) {
      setState(() {
        _status = '❌ STREAM CRASH with $rowCount rows!\n'
            'Error: $e\n'
            'Stack: ${stackTrace.toString().length > 300 ? stackTrace.toString().substring(0, 300) : stackTrace}';
      });
      if (kDebugMode) {
        print('Error streaming large Excel: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() => _isGenerating = false);
    }
  }

  Future<void> _streamReadExcel() async {
    setState(() {
      _isGenerating = true;
      _status = '📂 Pick an XLSX file to stream-read...';
    });

    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
        withData: true,
      );

      if (result == null || result.files.single.bytes == null) {
        setState(() {
          _isGenerating = false;
          _status = '⚠️ No file selected';
        });
        return;
      }

      final fileName = result.files.single.name;
      final bytes = result.files.single.bytes!;

      setState(() {
        _status = '📖 Parsing "$fileName" (${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB)...\n'
            '⏳ Loading shared strings...';
      });
      await Future.delayed(Duration.zero);

      final stopwatch = Stopwatch()..start();
      final reader = ExcelStreamReader.fromBytes(bytes);
      final parseTime = stopwatch.elapsed;

      final sheets = reader.sheetNames;
      final results = StringBuffer();
      results.writeln('✅ Stream read "$fileName" complete!');
      results.writeln('Archive parse: ${parseTime.inMilliseconds}ms');
      results.writeln('Sheets: ${sheets.join(", ")}');
      results.writeln();

      for (final sheetName in sheets) {
        setState(() {
          _status = '📖 Reading sheet "$sheetName"...';
        });
        await Future.delayed(Duration.zero);

        stopwatch.reset();
        var rowCount = 0;
        var cellCount = 0;

        for (final row in reader.readSheet(sheetName)) {
          rowCount++;
          cellCount += row.cells.length;

          if (rowCount % 10000 == 0) {
            setState(() {
              _status = '📖 Reading "$sheetName"...\n'
                  '📝 $rowCount rows read\n'
                  '⏱ Elapsed: ${stopwatch.elapsed.inSeconds}s';
            });
            await Future.delayed(Duration.zero);
          }
        }

        final readTime = stopwatch.elapsed;
        results.writeln('📄 $sheetName: $rowCount rows, $cellCount cells (${readTime.inMilliseconds}ms)');
      }

      setState(() {
        _status = results.toString();
      });
    } catch (e, stackTrace) {
      setState(() {
        _status = '❌ STREAM READ CRASH!\n'
            'Error: $e\n'
            'Stack: ${stackTrace.toString().length > 300 ? stackTrace.toString().substring(0, 300) : stackTrace}';
      });
      if (kDebugMode) {
        print('Error stream reading Excel: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() => _isGenerating = false);
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(20.0),
        child: Center(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.center,
            children: <Widget>[
              const Icon(Icons.table_chart, size: 80, color: Colors.blue),
              const SizedBox(height: 20),
              Text(
                'Excel Charts Demo',
                style: Theme.of(context).textTheme.headlineMedium,
              ),
              const SizedBox(height: 10),
              Text(
                _status,
                textAlign: TextAlign.center,
                style: const TextStyle(fontSize: 16),
              ),
              const SizedBox(height: 40),
              if (_isGenerating)
                const CircularProgressIndicator()
              else ...[
                ElevatedButton.icon(
                  onPressed: _generateSimpleExcel,
                  icon: const Icon(Icons.table_view),
                  label: const Text('Test: Simple Excel (No Chart)'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.green.shade100,
                  ),
                ),
                const SizedBox(height: 20),
                const Text(
                  'Excel with Charts:',
                  style: TextStyle(fontWeight: FontWeight.bold),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.column),
                  icon: const Icon(Icons.bar_chart),
                  label: const Text('Generate Column Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.line),
                  icon: const Icon(Icons.show_chart),
                  label: const Text('Generate Line Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.pie),
                  icon: const Icon(Icons.pie_chart),
                  label: const Text('Generate Pie Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.area),
                  icon: const Icon(Icons.area_chart),
                  label: const Text('Generate Area Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.doughnut),
                  icon: const Icon(Icons.donut_small),
                  label: const Text('Generate Doughnut Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.radar),
                  icon: const Icon(Icons.radar),
                  label: const Text('Generate Radar Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.bar),
                  icon: const Icon(Icons.horizontal_split),
                  label: const Text('Generate Bar Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.scatter),
                  icon: const Icon(Icons.scatter_plot),
                  label: const Text('Generate Scatter Chart'),
                ),
                const SizedBox(height: 30),
                const Divider(),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: _generateFullExample,
                  icon: const Icon(Icons.star, color: Colors.amber),
                  label: const Text('GENERATE FULL DEMO (Everything)'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.blue.shade50,
                    padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 15),
                    textStyle: const TextStyle(fontWeight: FontWeight.bold),
                  ),
                ),
                const SizedBox(height: 30),
                const Divider(),
                const SizedBox(height: 10),
                const Text(
                  '⚡ Large File Stress Tests:',
                  style: TextStyle(fontWeight: FontWeight.bold, color: Colors.red),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateLargeExcel(1000),
                  icon: const Icon(Icons.speed),
                  label: const Text('1,000 rows × 20 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.orange.shade100,
                  ),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateLargeExcel(5000),
                  icon: const Icon(Icons.speed),
                  label: const Text('5,000 rows × 20 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.orange.shade200,
                  ),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateLargeExcel(10000),
                  icon: const Icon(Icons.speed),
                  label: const Text('10,000 rows × 20 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.red.shade100,
                  ),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateLargeExcel(400000),
                  icon: const Icon(Icons.speed),
                  label: const Text('400,000 rows × 10 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.red.shade300,
                  ),
                ),
                const SizedBox(height: 30),
                const Divider(),
                const SizedBox(height: 10),
                const Text(
                  '🚀 Streaming API Stress Tests:',
                  style: TextStyle(fontWeight: FontWeight.bold, color: Colors.green),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _streamLargeExcel(10000),
                  icon: const Icon(Icons.stream),
                  label: const Text('Stream 10,000 rows × 20 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.green.shade100,
                  ),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _streamLargeExcel(100000),
                  icon: const Icon(Icons.stream),
                  label: const Text('Stream 100,000 rows × 20 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.green.shade200,
                  ),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _streamLargeExcel(400000),
                  icon: const Icon(Icons.stream),
                  label: const Text('Stream 400,000 rows × 20 cols'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.green.shade400,
                  ),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: _streamReadExcel,
                  icon: const Icon(Icons.file_open),
                  label: const Text('Stream Read XLSX File'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.blue.shade200,
                  ),
                ),
              ],
            ],
          ),
        ),
      ),
    );
  }
}

enum ChartType { column, line, pie, area, doughnut, radar, bar, scatter }

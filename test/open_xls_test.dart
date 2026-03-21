import 'package:dart_xls/dart_xls.dart';
import 'package:test/test.dart';

void main() {
  test('open sample xls file and parse heads', () {
    const path = 'test/file_example_XLS_50.xls';
    final wb = XlsWorkbook.openFile(path);
    expect(wb, isNotNull);

    wb.parse();
    final sheetCount = wb.sheetCount(maxProbe: 10);
    expect(sheetCount, greaterThan(0));

    final ws = wb.worksheet(0);
    ws.parse();
    final cells = ws.allCells(maxRows: 5, maxCols: 5);
    expect(cells, isNotEmpty);

    print('sheetCount=$sheetCount cells=${cells.length}');

    ws.close();
    wb.close();
  });
}

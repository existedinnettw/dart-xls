# dart_xls

Dart FFI bindings for libxls (XLS binary Excel parser).

## Setup

```bash
dart pub get
```

## Testing

```bash
dart pub get
dart test
```

## Usage

```dart
import 'dart:io';
import 'package:dart_xls/xls.dart';

void main() {
  final wb = XlsWorkbook.openFile('test/file_example_XLS_50.xls');
  wb.parse();

  final sheetCount = wb.sheetCount(maxProbe: 10);
  print('Sheet count: $sheetCount');

  for (var i = 0; i < sheetCount; i++) {
    final ws = wb.worksheet(i);
    ws.parse();

    final cells = ws.allCells(maxRows: 20, maxCols: 20);
    for (final cell in cells) {
      print('R${cell.row}C${cell.col}: ${cell.text ?? cell.numericValue}');
    }

    ws.close();
  }

  wb.close();
}
```

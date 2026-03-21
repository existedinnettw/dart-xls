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

## Platforms

- linux: supported via native `libxls` (either submodule build or system package).
- macos: supported via native `libxls` (.dylib).
- windows: supported via native `libxls` (.dll).
- android/ios: supported via Flutter embedding if `libxls` is bundled as a native library for the platform.
- web: unsupported (Dart web does not support dart:ffi, `web` is marked unsupported in `pubspec.yaml`).

# TODO
- [ ] web platform support

# references

- [libxls](https://github.com/libxls/libxls)
- [sqlite3.dart](https://github.com/simolus3/sqlite3.dart/tree/main)
  - The multiplatform FFI bindings and native library loading patterns in `sqlite3.dart` were a helpful reference for this project.

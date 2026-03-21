/// Dart FFI binding for libxls (https://github.com/libxls/libxls)
///
/// Requires the native libxls shared library installed on the host system
/// (libxls.so / libxls.dylib / libxls.dll).

library xls;

import 'dart:ffi';
import 'dart:io';
import 'dart:typed_data';

import 'package:ffi/ffi.dart';
import 'package:path/path.dart' as p;

// API error codes from xls.h
enum XlsError {
  ok,
  error,
  open,
  seek,
  read,
  parse,
  malloc,
  unsupportedEncryption,
  nullArgument,
}

class XlsException implements Exception {
  final XlsError error;
  final String message;

  XlsException(this.error, this.message);

  @override
  String toString() => 'XlsException($error): $message';
}

class _LibXls {
  static final DynamicLibrary _dl = _openLibrary();

  static String? get _repoRoot {
    final current = Directory.current.path;
    if (Directory(p.join(current, 'libxls')).existsSync()) {
      return current;
    }
    final script = p.dirname(Platform.script.toFilePath());
    if (Directory(p.join(script, 'libxls')).existsSync()) {
      return script;
    }
    return null;
  }

  static List<String> get _submoduleCandidates {
    final root = _repoRoot;
    if (root == null) return [];
    final base = p.join(root, 'libxls');
    return [
      p.join(base, '.libs', 'libxls.so'),
      p.join(base, '.libs', 'libxls.dylib'),
      p.join(base, '.libs', 'libxls.dll'),
      p.join(base, '.libs', 'libxlsreader.so'),
      p.join(base, '.libs', 'libxlsreader.so.8'),
      p.join(base, '.libs', 'libxlsreader.so.8.0.3'),
      p.join(base, '.libs', 'libxlsreader.dylib'),
      p.join(base, '.libs', 'libxlsreader.dll'),
      p.join(base, 'src', '.libs', 'libxls.so'),
      p.join(base, 'src', '.libs', 'libxls.dylib'),
    ];
  }

  static void _tryBuildLibxls() {
    final root = _repoRoot;
    if (root == null) return;
    final sub = Directory(p.join(root, 'libxls'));
    if (!sub.existsSync()) return;

    try {
      if (!File(p.join(sub.path, 'configure')).existsSync()) {
        final a1 = Process.runSync('./autogen.sh', [],
            workingDirectory: sub.path, runInShell: true);
        if (a1.exitCode != 0) return;
      }

      final conf = Process.runSync('./configure', [],
          workingDirectory: sub.path, runInShell: true);
      if (conf.exitCode != 0) return;

      final m = Process.runSync('make', ['-j', '4'],
          workingDirectory: sub.path, runInShell: true);
      if (m.exitCode != 0) return;
    } catch (_) {
      // ignore build failure; fallback to expected system path
    }
  }

  static List<String> get _systemLibCandidates {
    if (Platform.isLinux) {
      return ['libxls.so', 'libxlsreader.so'];
    } else if (Platform.isMacOS) {
      return ['libxls.dylib', 'libxlsreader.dylib'];
    } else if (Platform.isWindows) {
      return ['libxls.dll', 'libxlsreader.dll'];
    }
    return [];
  }

  static DynamicLibrary _openLibrary() {
    final candidates = <String>[];
    candidates.addAll(_submoduleCandidates);

    for (final candidate in candidates) {
      if (File(candidate).existsSync()) {
        return DynamicLibrary.open(candidate);
      }
    }

    _tryBuildLibxls();

    // Re-check after build attempt
    for (final candidate in _submoduleCandidates) {
      if (File(candidate).existsSync()) {
        return DynamicLibrary.open(candidate);
      }
    }

    // Fallback to system-installed libxls
    for (final candidate in _systemLibCandidates) {
      try {
        return DynamicLibrary.open(candidate);
      } catch (_) {
        // ignore and continue
      }
    }

    throw StateError(
        'Cannot find libxls in libxls submodule, system library path, and build failed.\n'
        'Install libxls from your package manager or build it from source.');
  }

  static final _XlsFunctions _f = _XlsFunctions(_dl);

  static _XlsFunctions get f => _f;
}

class _XlsFunctions {
  _XlsFunctions(DynamicLibrary dl) {
    xlsGetVersion =
        dl.lookupFunction<Pointer<Utf8> Function(), Pointer<Utf8> Function()>(
            'xls_getVersion');
    xlsGetError = dl.lookupFunction<Pointer<Utf8> Function(Int32),
        Pointer<Utf8> Function(int)>('xls_getError');
    xlsOpenFile = dl.lookupFunction<
        Pointer<Void> Function(Pointer<Int8>, Pointer<Int8>, Pointer<Int32>),
        Pointer<Void> Function(
            Pointer<Int8>, Pointer<Int8>, Pointer<Int32>)>('xls_open_file');
    xlsOpenBuffer = dl.lookupFunction<
        Pointer<Void> Function(
            Pointer<Uint8>, IntPtr, Pointer<Int8>, Pointer<Int32>),
        Pointer<Void> Function(Pointer<Uint8>, int, Pointer<Int8>,
            Pointer<Int32>)>('xls_open_buffer');
    xlsParseWorkBook = dl.lookupFunction<Int32 Function(Pointer<Void>),
        int Function(Pointer<Void>)>('xls_parseWorkBook');
    xlsParseWorkSheet = dl.lookupFunction<Int32 Function(Pointer<Void>),
        int Function(Pointer<Void>)>('xls_parseWorkSheet');
    xlsCloseWB = dl.lookupFunction<Void Function(Pointer<Void>),
        void Function(Pointer<Void>)>('xls_close_WB');
    xlsCloseWS = dl.lookupFunction<Void Function(Pointer<Void>),
        void Function(Pointer<Void>)>('xls_close_WS');
    xlsGetWorkSheet = dl.lookupFunction<
        Pointer<Void> Function(Pointer<Void>, Int32),
        Pointer<Void> Function(Pointer<Void>, int)>('xls_getWorkSheet');
    xlsRow = dl.lookupFunction<
        Pointer<XlsRowStruct> Function(Pointer<Void>, Uint16),
        Pointer<XlsRowStruct> Function(Pointer<Void>, int)>('xls_row');
    xlsCell = dl.lookupFunction<
        Pointer<XlsCell> Function(Pointer<Void>, Uint16, Uint16),
        Pointer<XlsCell> Function(Pointer<Void>, int, int)>('xls_cell');
  }

  late final Pointer<Utf8> Function() xlsGetVersion;
  late final Pointer<Utf8> Function(int) xlsGetError;

  late final Pointer<Void> Function(
      Pointer<Int8>, Pointer<Int8>, Pointer<Int32>) xlsOpenFile;
  late final Pointer<Void> Function(
      Pointer<Uint8>, int, Pointer<Int8>, Pointer<Int32>) xlsOpenBuffer;
  late final int Function(Pointer<Void>) xlsParseWorkBook;
  late final int Function(Pointer<Void>) xlsParseWorkSheet;
  late final void Function(Pointer<Void>) xlsCloseWB;
  late final void Function(Pointer<Void>) xlsCloseWS;
  late final Pointer<Void> Function(Pointer<Void>, int) xlsGetWorkSheet;
  late final Pointer<XlsRowStruct> Function(Pointer<Void>, int) xlsRow;
  late final Pointer<XlsCell> Function(Pointer<Void>, int, int) xlsCell;
}

class XlsCell extends Struct {
  @Uint16()
  external int id;

  @Uint16()
  external int row;

  @Uint16()
  external int col;

  @Uint16()
  external int xf;

  external Pointer<Utf8> str;

  @Double()
  external double d;

  @Int32()
  external int l;

  @Uint16()
  external int width;

  @Uint16()
  external int colspan;

  @Uint16()
  external int rowspan;

  @Uint8()
  external int isHidden;
}

class XlsRowStruct extends Struct {
  @Uint16()
  external int index;

  @Uint16()
  external int fcell;

  @Uint16()
  external int lcell;

  @Uint16()
  external int height;

  @Uint16()
  external int flags;

  @Uint16()
  external int xf;

  @Uint8()
  external int xfflags;

  external Pointer<XlsCell> cells;
}

class XlsWorkbook {
  final Pointer<Void> _ptr;

  XlsWorkbook._(this._ptr);

  static XlsWorkbook openFile(String path, {String charset = 'UTF-8'}) {
    final outError = calloc<Int32>();
    final libPath = path.toNativeUtf8().cast<Int8>();
    final libCharset = charset.toNativeUtf8().cast<Int8>();
    final wbPtr = _LibXls.f.xlsOpenFile(libPath, libCharset, outError);

    calloc.free(libPath);
    calloc.free(libCharset);

    if (wbPtr == nullptr) {
      final err = outError.value;
      final codeIndex = err.clamp(0, XlsError.values.length - 1);
      final code = XlsError.values[codeIndex];
      final message = _getErrorText(err);
      calloc.free(outError);
      throw XlsException(code, 'xlsOpenFile failed: $message');
    }

    final err = outError.value;
    calloc.free(outError);
    if (err != 0) {
      final codeIndex = err.clamp(0, XlsError.values.length - 1);
      final code = XlsError.values[codeIndex];
      final message = _getErrorText(err);
      throw XlsException(code, 'xlsOpenFile parse error: $message');
    }

    return XlsWorkbook._(wbPtr);
  }

  static XlsWorkbook openBuffer(Uint8List bytes, {String charset = 'UTF-8'}) {
    final outError = calloc<Int32>();
    final data = calloc<Uint8>(bytes.length);
    final libCharset = charset.toNativeUtf8().cast<Int8>();
    final ptr = data;
    final list = ptr.asTypedList(bytes.length);
    list.setAll(0, bytes);

    final wbPtr = _LibXls.f
        .xlsOpenBuffer(ptr, bytes.length, libCharset.cast<Int8>(), outError);

    calloc.free(data);
    calloc.free(libCharset);

    if (wbPtr == nullptr) {
      final err = outError.value;
      final codeIndex = err.clamp(0, XlsError.values.length - 1);
      final code = XlsError.values[codeIndex];
      final message = _getErrorText(err);
      calloc.free(outError);
      throw XlsException(code, 'xlsOpenBuffer failed: $message');
    }

    final err = outError.value;
    calloc.free(outError);
    if (err != 0) {
      final codeIndex = err.clamp(0, XlsError.values.length - 1);
      final code = XlsError.values[codeIndex];
      final message = _getErrorText(err);
      throw XlsException(code, 'xlsOpenBuffer parse error: $message');
    }

    return XlsWorkbook._(wbPtr);
  }

  void parse() {
    final err = _LibXls.f.xlsParseWorkBook(_ptr);
    if (err != 0) {
      throw XlsException(
          XlsError.values[err.clamp(0, XlsError.values.length - 1)],
          _getErrorText(err));
    }
  }

  XlsWorksheet worksheet(int index) {
    final wsPtr = _LibXls.f.xlsGetWorkSheet(_ptr, index);
    if (wsPtr == nullptr) {
      throw RangeError.index(
          index, null, 'index', 'Worksheet index out of range');
    }
    return XlsWorksheet._(wsPtr);
  }

  int sheetCount({int maxProbe = 512}) {
    for (var i = 0; i < maxProbe; ++i) {
      final ws = _LibXls.f.xlsGetWorkSheet(_ptr, i);
      if (ws == nullptr) {
        return i;
      }
      _LibXls.f.xlsCloseWS(ws); // close immediately to avoid leaks in probe
    }
    return maxProbe;
  }

  void close() {
    _LibXls.f.xlsCloseWB(_ptr);
  }
}

class XlsWorksheet {
  final Pointer<Void> _ptr;

  XlsWorksheet._(this._ptr);

  void parse() {
    final err = _LibXls.f.xlsParseWorkSheet(_ptr);
    if (err != 0) {
      throw XlsException(
          XlsError.values[err.clamp(0, XlsError.values.length - 1)],
          _getErrorText(err));
    }
  }

  XlsRowData? row(int rowIndex) {
    final rowPtr = _LibXls.f.xlsRow(_ptr, rowIndex);
    if (rowPtr == nullptr) return null;
    return XlsRowData.fromPointer(rowPtr);
  }

  XlsCellData? cell(int rowIndex, int colIndex) {
    final cellPtr = _LibXls.f.xlsCell(_ptr, rowIndex, colIndex);
    if (cellPtr == nullptr) return null;
    return XlsCellData.fromPointer(cellPtr.cast<XlsCell>());
  }

  List<XlsCellData> allCells({int maxRows = 1024, int maxCols = 1024}) {
    final cells = <XlsCellData>[];
    for (var r = 0; r <= maxRows; ++r) {
      final rowData = row(r);
      if (rowData == null) break;
      if (rowData.firstCell > rowData.lastCell) continue;
      for (var c = rowData.firstCell; c <= rowData.lastCell; ++c) {
        final xlsCell = cell(r, c);
        if (xlsCell != null) cells.add(xlsCell);
      }
    }
    return cells;
  }

  void close() {
    _LibXls.f.xlsCloseWS(_ptr);
  }
}

class XlsRowData {
  final int index;
  final int firstCell;
  final int lastCell;
  final int height;
  final int flags;
  final int xf;
  final Pointer<XlsCell> cells;

  XlsRowData({
    required this.index,
    required this.firstCell,
    required this.lastCell,
    required this.height,
    required this.flags,
    required this.xf,
    required this.cells,
  });

  factory XlsRowData.fromPointer(Pointer<XlsRowStruct> ptr) {
    final row = ptr.ref;
    return XlsRowData(
      index: row.index,
      firstCell: row.fcell,
      lastCell: row.lcell,
      height: row.height,
      flags: row.flags,
      xf: row.xf,
      cells: row.cells,
    );
  }
}

class XlsCellData {
  final int id;
  final int row;
  final int col;
  final int xf;
  final String? text;
  final double numericValue;
  final int longValue;
  final int width;
  final int colspan;
  final int rowspan;
  final bool isHidden;

  XlsCellData({
    required this.id,
    required this.row,
    required this.col,
    required this.xf,
    this.text,
    required this.numericValue,
    required this.longValue,
    required this.width,
    required this.colspan,
    required this.rowspan,
    required this.isHidden,
  });

  factory XlsCellData.fromPointer(Pointer<XlsCell> ptr) {
    final cell = ptr.ref;
    final text =
        cell.str == nullptr ? null : Utf8Pointer(cell.str).toDartString();
    return XlsCellData(
      id: cell.id,
      row: cell.row,
      col: cell.col,
      xf: cell.xf,
      text: text,
      numericValue: cell.d,
      longValue: cell.l,
      width: cell.width,
      colspan: cell.colspan,
      rowspan: cell.rowspan,
      isHidden: cell.isHidden != 0,
    );
  }

  bool get isBlank => id == 0x0201;
  bool get isNumber => id == 0x0203;
  bool get isLabel => id == 0x0204;
  bool get isBooleanOrError => id == 0x0205;
  bool get isString => id == 0x0207;
  bool get isFormula => id == 0x0006;
}

String _getErrorText(int code) {
  if (code >= 0 && code < XlsError.values.length) {
    final ptr = _LibXls.f.xlsGetError(code);
    if (ptr != nullptr) {
      return Utf8Pointer(ptr).toDartString();
    }
  }

  return 'Unknown libxls error code $code';
}

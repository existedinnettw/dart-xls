import 'dart:ffi';

import 'package:dart_xls/xls.dart';
import 'package:ffi/ffi.dart';
import 'package:test/test.dart';

void main() {
  group('XlsError enum', () {
    test('values are ordered as in c API', () {
      expect(XlsError.ok.index, equals(0));
      expect(XlsError.open.index, equals(2));
      expect(XlsError.nullArgument.index, equals(8));
      expect(XlsError.values.length, greaterThan(0));
    });
  });

  group('XlsCellData.fromPointer', () {
    test('reads field values correctly', () {
      final ptr = calloc<XlsCell>();
      try {
        final cell = ptr.ref;
        cell.id = 0x0204;
        cell.row = 1;
        cell.col = 2;
        cell.xf = 0;
        cell.str = 'hello'.toNativeUtf8();
        cell.d = 3.1415;
        cell.l = 123;
        cell.width = 10;
        cell.colspan = 1;
        cell.rowspan = 1;
        cell.isHidden = 0;

        final dto = XlsCellData.fromPointer(ptr);
        expect(dto.id, equals(0x0204));
        expect(dto.row, equals(1));
        expect(dto.col, equals(2));
        expect(dto.text, equals('hello'));
        expect(dto.numericValue, closeTo(3.1415, 1e-12));
        expect(dto.longValue, equals(123));
        expect(dto.width, equals(10));
        expect(dto.colspan, equals(1));
        expect(dto.rowspan, equals(1));
        expect(dto.isHidden, isFalse);
      } finally {
        calloc.free(ptr.ref.str as Pointer<Void>);
        calloc.free(ptr);
      }
    });
  });

  group('XlsRowData.fromPointer', () {
    test('reads row metadata and pointer', () {
      final rowPtr = calloc<XlsRowStruct>();
      try {
        final row = rowPtr.ref;
        row.index = 5;
        row.fcell = 0;
        row.lcell = 3;
        row.height = 256;
        row.flags = 1;
        row.xf = 0;
        row.xfflags = 0;
        row.cells = nullptr; // optional

        final dto = XlsRowData.fromPointer(rowPtr);
        expect(dto.index, equals(5));
        expect(dto.firstCell, equals(0));
        expect(dto.lastCell, equals(3));
        expect(dto.height, equals(256));
        expect(dto.flags, equals(1));
        expect(dto.xf, equals(0));
        expect(dto.cells, equals(nullptr));
      } finally {
        calloc.free(rowPtr);
      }
    });
  });
}

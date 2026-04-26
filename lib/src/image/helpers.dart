// docx/image/helpers.dart
import 'dart:convert';
import 'dart:typed_data';

import 'exceptions.dart';

// Constantes para Endian podem ser usadas diretamente de dart:typed_data
// final Endian BIG_ENDIAN = Endian.big;
// final Endian LITTLE_ENDIAN = Endian.little;

class StreamReader {
  final ByteData _byteData;
  final Endian _endian;
  int _offset = 0; // Posição atual de leitura

  StreamReader(Uint8List bytes, this._endian)
      : _byteData = ByteData.view(bytes.buffer);

  // Métodos adaptados para usar ByteData
  Uint8List read(int count) {
    final result =
        _byteData.buffer.asUint8List(_byteData.offsetInBytes + _offset, count);
    _offset += count;
    return result;
  }

  int readByte(int base, {int offset = 0}) {
    seek(base, offset: offset);
    final val = _byteData.getUint8(_offset);
    _offset += 1;
    return val;
  }

  int readLong(int base, {int offset = 0}) {
    seek(base, offset: offset);
    final val = _byteData.getUint32(_offset, _endian);
    _offset += 4;
    return val;
  }

  int readShort(int base, {int offset = 0}) {
    seek(base, offset: offset);
    final val = _byteData.getUint16(_offset, _endian);
    _offset += 2;
    return val;
  }

  String readString(int charCount, int base, {int offset = 0}) {
    final bytes = _readBytes(charCount, base, offset: offset);
    return utf8.decode(bytes);
  }

  void seek(int base, {int offset = 0}) {
    final location = base + offset;
    if (location < 0 || location >= _byteData.lengthInBytes) {
      throw RangeError('Seek out of bounds');
    }
    _offset = location;
  }

  int tell() {
    return _offset;
  }

  Uint8List _readBytes(int byteCount, int base, {int offset = 0}) {
    seek(base, offset: offset);
    if (_offset + byteCount > _byteData.lengthInBytes) {
      throw UnexpectedEndOfFileError('unexpected end of file');
    }
    final result = _byteData.buffer.asUint8List(
      _byteData.offsetInBytes + _offset,
      byteCount,
    );
    _offset += byteCount;
    return result;
  }
}

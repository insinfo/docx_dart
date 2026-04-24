import 'dart:convert';
import 'dart:typed_data';

import 'package:docx_dart/src/internal/archive/codecs/zlib/deflate.dart';
import 'package:docx_dart/src/internal/archive/codecs/zlib/inflate.dart';
import 'package:docx_dart/src/internal/archive/util/output_memory_stream.dart';
import 'package:docx_dart/src/internal/archive/util/crc32.dart';

const int _zipLocalFileHeaderSignature = 0x04034b50;
const int _zipCentralDirectorySignature = 0x02014b50;
const int _zipEndOfCentralDirectorySignature = 0x06054b50;
const int _zipCompressionStored = 0;
const int _zipCompressionDeflate = 8;
const int _zipUtf8Flag = 0x0800;

class ZipArchiveEntry {
  final String name;
  final Uint8List content;

  ZipArchiveEntry(this.name, List<int> content)
      : content = content is Uint8List ? content : Uint8List.fromList(content);
}

class ZipArchive {
  final List<ZipArchiveEntry> _entries = [];
  final Map<String, int> _entryIndex = <String, int>{};

  ZipArchive();

  List<ZipArchiveEntry> get entries => List<ZipArchiveEntry>.unmodifiable(_entries);

  ZipArchiveEntry? findFile(String name) {
    final index = _entryIndex[name];
    return index == null ? null : _entries[index];
  }

  void addFile(String name, List<int> content) {
    final entry = ZipArchiveEntry(name, content);
    final existingIndex = _entryIndex[name];
    if (existingIndex != null) {
      _entries[existingIndex] = entry;
      return;
    }
    _entryIndex[name] = _entries.length;
    _entries.add(entry);
  }

  factory ZipArchive.decodeBytes(Uint8List bytes) {
    final archive = ZipArchive();
    if (bytes.isEmpty) {
      return archive;
    }

    final eocdOffset = _findEndOfCentralDirectory(bytes);
    if (eocdOffset < 0) {
      throw FormatException('Invalid ZIP archive: end of central directory not found.');
    }

    final entryCount = _readUint16(bytes, eocdOffset + 10);
    final centralDirectorySize = _readUint32(bytes, eocdOffset + 12);
    final centralDirectoryOffset = _readUint32(bytes, eocdOffset + 16);
    final centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;

    var offset = centralDirectoryOffset;
    for (var index = 0; index < entryCount && offset < centralDirectoryEnd; index++) {
      final signature = _readUint32(bytes, offset);
      if (signature != _zipCentralDirectorySignature) {
        throw FormatException('Invalid ZIP archive: unexpected central directory header.');
      }

      final flags = _readUint16(bytes, offset + 8);
      final compressionMethod = _readUint16(bytes, offset + 10);
      final compressedSize = _readUint32(bytes, offset + 20);
      final uncompressedSize = _readUint32(bytes, offset + 24);
      final fileNameLength = _readUint16(bytes, offset + 28);
      final extraFieldLength = _readUint16(bytes, offset + 30);
      final commentLength = _readUint16(bytes, offset + 32);
      final localHeaderOffset = _readUint32(bytes, offset + 42);
      final nameStart = offset + 46;
      final nameEnd = nameStart + fileNameLength;
      final fileName = _decodeName(
        bytes,
        nameStart,
        nameEnd,
        useUtf8: (flags & _zipUtf8Flag) != 0,
      );

      if (compressedSize == 0xffffffff || uncompressedSize == 0xffffffff) {
        throw UnsupportedError('ZIP64 archives are not supported.');
      }

      final localSignature = _readUint32(bytes, localHeaderOffset);
      if (localSignature != _zipLocalFileHeaderSignature) {
        throw FormatException('Invalid ZIP archive: local file header not found.');
      }

      final localFileNameLength = _readUint16(bytes, localHeaderOffset + 26);
      final localExtraFieldLength = _readUint16(bytes, localHeaderOffset + 28);
      final dataStart = localHeaderOffset + 30 + localFileNameLength + localExtraFieldLength;
      final dataEnd = dataStart + compressedSize;
      final compressedBytes = Uint8List.sublistView(bytes, dataStart, dataEnd);
      final content = _decodeContent(compressedBytes, compressionMethod, uncompressedSize);

      archive.addFile(fileName, content);
      offset = nameEnd + extraFieldLength + commentLength;
    }

    return archive;
  }

  Uint8List encode() {
    final output = OutputMemoryStream();
    final centralDirectory = OutputMemoryStream();
    final now = DateTime.now();
    final dosTime = _dosTime(now);
    final dosDate = _dosDate(now);

    for (final entry in _entries) {
      final nameBytes = Uint8List.fromList(utf8.encode(entry.name));
      final compressedBytes = _compress(entry.content);
      final useCompression = compressedBytes.length < entry.content.length;
      final method = useCompression ? _zipCompressionDeflate : _zipCompressionStored;
      final payload = useCompression ? compressedBytes : entry.content;
      final crc32 = getCrc32(entry.content);
      final localHeaderOffset = output.length;

      output
        ..writeUint32(_zipLocalFileHeaderSignature)
        ..writeUint16(20)
        ..writeUint16(_zipUtf8Flag)
        ..writeUint16(method)
        ..writeUint16(dosTime)
        ..writeUint16(dosDate)
        ..writeUint32(crc32)
        ..writeUint32(payload.length)
        ..writeUint32(entry.content.length)
        ..writeUint16(nameBytes.length)
        ..writeUint16(0)
        ..writeBytes(nameBytes)
        ..writeBytes(payload);

      centralDirectory
        ..writeUint32(_zipCentralDirectorySignature)
        ..writeUint16(20)
        ..writeUint16(20)
        ..writeUint16(_zipUtf8Flag)
        ..writeUint16(method)
        ..writeUint16(dosTime)
        ..writeUint16(dosDate)
        ..writeUint32(crc32)
        ..writeUint32(payload.length)
        ..writeUint32(entry.content.length)
        ..writeUint16(nameBytes.length)
        ..writeUint16(0)
        ..writeUint16(0)
        ..writeUint16(0)
        ..writeUint16(0)
        ..writeUint32(0)
        ..writeUint32(localHeaderOffset)
        ..writeBytes(nameBytes);
    }

    final centralDirectoryOffset = output.length;
    final centralDirectoryBytes = centralDirectory.getBytes();
    output.writeBytes(centralDirectoryBytes);

    output
      ..writeUint32(_zipEndOfCentralDirectorySignature)
      ..writeUint16(0)
      ..writeUint16(0)
      ..writeUint16(_entries.length)
      ..writeUint16(_entries.length)
      ..writeUint32(centralDirectoryBytes.length)
      ..writeUint32(centralDirectoryOffset)
      ..writeUint16(0);

    return output.getBytes();
  }
}

Uint8List _compress(Uint8List bytes) {
  final output = OutputMemoryStream();
  Deflate(bytes, output: output);
  return output.getBytes();
}

Uint8List _decodeContent(Uint8List compressedBytes, int method, int uncompressedSize) {
  if (method == _zipCompressionStored) {
    return Uint8List.fromList(compressedBytes);
  }
  if (method == _zipCompressionDeflate) {
    return Inflate(compressedBytes, uncompressedSize: uncompressedSize).getBytes();
  }
  throw UnsupportedError('ZIP compression method $method is not supported.');
}

int _findEndOfCentralDirectory(Uint8List bytes) {
  final lowerBound = bytes.length > 0x10016 ? bytes.length - 0x10016 : 0;
  for (var offset = bytes.length - 22; offset >= lowerBound; offset--) {
    if (_readUint32(bytes, offset) == _zipEndOfCentralDirectorySignature) {
      return offset;
    }
  }
  return -1;
}

String _decodeName(Uint8List bytes, int start, int end, {required bool useUtf8}) {
  final slice = Uint8List.sublistView(bytes, start, end);
  if (!useUtf8) {
    return latin1.decode(slice);
  }
  try {
    return utf8.decode(slice);
  } catch (_) {
    return latin1.decode(slice);
  }
}

int _readUint16(Uint8List bytes, int offset) =>
    bytes[offset] | (bytes[offset + 1] << 8);

int _readUint32(Uint8List bytes, int offset) =>
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24);

int _dosTime(DateTime dateTime) =>
    ((dateTime.hour & 0x1f) << 11) |
    ((dateTime.minute & 0x3f) << 5) |
    ((dateTime.second ~/ 2) & 0x1f);

int _dosDate(DateTime dateTime) =>
    (((dateTime.year - 1980) & 0x7f) << 9) |
    ((dateTime.month & 0xf) << 5) |
    (dateTime.day & 0x1f);

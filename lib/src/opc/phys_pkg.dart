// docx/opc/phys_pkg.dart
import 'dart:typed_data';
import 'package:docx_dart/src/opc/exceptions.dart';
import 'package:docx_dart/src/opc/internal_zip.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/platform/file_access.dart';

abstract class PhysPkgReader {
  factory PhysPkgReader(dynamic pkgFile) {
    if (pkgFile is String) {
      if (!supportsFileAccess) {
        throw UnsupportedError(
          'Filesystem path access is not supported on this platform. Pass ZIP bytes instead.',
        );
      }
      if (isDirectoryPathSync(pkgFile)) {
        throw UnimplementedError("Directory package reading not supported");
      } else if (isFilePathSync(pkgFile)) {
        final bytes = readFileBytesSync(pkgFile);
        if (bytes == null) {
          throw PackageNotFoundError("Package not found at '$pkgFile'");
        }
        return _ZipPkgReader.fromBytes(bytes);
      } else {
        throw PackageNotFoundError("Package not found at '$pkgFile'");
      }
    } else if (pkgFile is Uint8List) {
      return _ZipPkgReader.fromBytes(pkgFile);
    } else if (pkgFile is List<int>) {
      return _ZipPkgReader.fromBytes(Uint8List.fromList(pkgFile));
    } else {
      throw ArgumentError("Unsupported pkgFile type: ${pkgFile.runtimeType}");
    }
  }

  Uint8List blobFor(PackUri packUri);
  String? relsXmlFor(PackUri sourceUri);
  Uint8List get contentTypesXml;
  void close();
}

abstract class PhysPkgWriter {
  factory PhysPkgWriter(dynamic pkgFile) {
    if (pkgFile is String) {
      return _ZipPkgWriter(pkgFile);
    }
    if (pkgFile is BytesBuilder) {
      return _MemoryZipPkgWriter(pkgFile);
    }
    throw ArgumentError(
      'PhysPkgWriter supports either a String path or a BytesBuilder sink.',
    );
  }

  void write(PackUri packUri, Uint8List blob);
  void close();
}

class _ZipPkgReader implements PhysPkgReader {
  final ZipArchive _archive;

  _ZipPkgReader.fromBytes(Uint8List bytes)
      : _archive = ZipArchive.decodeBytes(bytes);

  @override
  Uint8List blobFor(PackUri packUri) {
    final file = _archive.findFile(packUri.membername);
    if (file == null) {
      throw PackageNotFoundError("Part not found: ${packUri.membername}");
    }
    return Uint8List.fromList(file.content);
  }

  @override
  String? relsXmlFor(PackUri sourceUri) {
    final file = _archive.findFile(sourceUri.relsUri.membername);
    if (file == null) return null;
    return String.fromCharCodes(file.content);
  }

  @override
  Uint8List get contentTypesXml => blobFor(CONTENT_TYPES_URI);

  @override
  void close() {}
}

class _ZipPkgWriter implements PhysPkgWriter {
  final String _path;
  final ZipArchive _archive = ZipArchive();

  _ZipPkgWriter(this._path);

  @override
  void write(PackUri packUri, Uint8List blob) {
    _archive.addFile(packUri.membername, blob);
  }

  @override
  void close() {
    writeFileBytesSync(_path, _archive.encode());
  }
}

class _MemoryZipPkgWriter implements PhysPkgWriter {
  final BytesBuilder _bytesBuilder;
  final ZipArchive _archive = ZipArchive();

  _MemoryZipPkgWriter(this._bytesBuilder);

  @override
  void write(PackUri packUri, Uint8List blob) {
    _archive.addFile(packUri.membername, blob);
  }

  @override
  void close() {
    _bytesBuilder.add(_archive.encode());
  }
}

// docx/opc/phys_pkg.dart
import 'dart:typed_data';
import 'dart:io';
import 'package:archive/archive.dart'; // Para manipulação de Zip
import 'package:docx_dart/src/opc/exceptions.dart';
import 'package:docx_dart/src/opc/packuri.dart';

abstract class PhysPkgReader {
  factory PhysPkgReader(dynamic pkgFile) {
    if (pkgFile is String) {
      if (FileSystemEntity.isDirectorySync(pkgFile)) {
        // return _DirPkgReader(pkgFile); // Leitura de diretório não implementada com 'archive'
        throw UnimplementedError("Directory package reading not supported");
      } else if (FileSystemEntity.isFileSync(pkgFile)) {
        // Verificar se é zip pode ser feito lendo bytes, mas `archive` faz isso.
        return _ZipPkgReader(pkgFile);
      } else {
        throw PackageNotFoundError("Package not found at '$pkgFile'");
      }
    } else if (pkgFile is Uint8List) {
      return _ZipPkgReader.fromBytes(pkgFile);
    } else if (pkgFile is List<int>) {
      return _ZipPkgReader.fromBytes(Uint8List.fromList(pkgFile));
    } else if (pkgFile is Archive) {
      return _ZipPkgReader.fromArchive(pkgFile); // Construtor adicional útil
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
     // Poderia adicionar suporte para escrever em um Sink<List<int>>
     throw ArgumentError("Only String path is supported for PhysPkgWriter currently");
   }

   void write(PackUri packUri, Uint8List blob);
   void close();
}


// Implementação Zip usando package:archive
class _ZipPkgReader implements PhysPkgReader {
  final Archive _archive;
  RandomAccessFile? _fileHandle; // Para fechar se abrimos o arquivo

  _ZipPkgReader(String path) :
     _archive = ZipDecoder().decodeBytes(File(path).readAsBytesSync()),
     _fileHandle = null;


   _ZipPkgReader.fromBytes(Uint8List bytes) :
     _archive = ZipDecoder().decodeBytes(bytes),
     _fileHandle = null;

   _ZipPkgReader.fromArchive(this._archive) :
      _fileHandle = null;


  @override
  Uint8List blobFor(PackUri packUri) {
    final file = _archive.findFile(packUri.membername);
    if (file == null) {
      throw PackageNotFoundError("Part not found: ${packUri.membername}");
    }
     // file.content pode ser dynamic (Uint8List ou List<int>), converter se necessário
    final List<int> content = file.content as List<int>;
    return Uint8List.fromList(List<int>.from(content));
  }

  @override
  String? relsXmlFor(PackUri sourceUri) {
    final file = _archive.findFile(sourceUri.relsUri.membername);
    if (file == null) return null;
     // Assumindo que rels são UTF8
          final List<int> content = file.content as List<int>;
          return String.fromCharCodes(content);
  }

  @override
  Uint8List get contentTypesXml => blobFor(CONTENT_TYPES_URI);

  @override
  void close() {
     _fileHandle?.closeSync(); // Fecha o arquivo se foi aberto por nós
  }
}


class _ZipPkgWriter implements PhysPkgWriter {
  final String _path;
  final Archive _archive = Archive(); // Começa com um arquivo vazio

  _ZipPkgWriter(this._path);

  @override
  void write(PackUri packUri, Uint8List blob) {
     final archiveFile = ArchiveFile(packUri.membername, blob.length, blob);
     _archive.addFile(archiveFile);
  }

  @override
  void close() {
      final encodedBytes = ZipEncoder().encode(_archive);
      File(_path).writeAsBytesSync(encodedBytes);
  }
}
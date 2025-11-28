// docx/image/image.dart
import 'dart:typed_data';
import 'dart:io';
import 'package:path/path.dart' as p; // Para basename e splitext
import 'package:crypto/crypto.dart' as crypto; // Para sha1
import 'package:docx_dart/src/shared.dart'; // Para Length, Inches, Emu

abstract class BaseImageHeader {
 final int pxWidth;
 final int pxHeight;
 final int horzDpi;
 final int vertDpi;

 BaseImageHeader(this.pxWidth, this.pxHeight, this.horzDpi, this.vertDpi);

 String get contentType; // Abstrato
 String get defaultExt; // Abstrato
}

class Image {
 final Uint8List _blob;
 final String _filename;
 final BaseImageHeader _imageHeader;

 Image._internal(this._blob, this._filename, this._imageHeader);

 /// Cria uma instância de [Image] a partir de um [Uint8List] (blob).
 static Image fromBlob(Uint8List blob) {
    final header = _ImageHeaderFactory(blob); // Supõe que a factory use o blob
    final filename = 'image.${header.defaultExt}';
    return Image._internal(blob, filename, header);
 }

 /// Cria uma instância de [Image] a partir de um caminho de arquivo.
 static Future<Image> fromPath(String imagePath) async {
    final file = File(imagePath);
    final blob = await file.readAsBytes();
    final header = _ImageHeaderFactory(blob);
    final filename = p.basename(imagePath);
     return Image._internal(blob, filename, header);
 }

 /// Cria uma instância de [Image] a partir de bytes, com um nome de arquivo opcional.
 static Image fromBytes(Uint8List blob, {String? filename}) {
    final header = _ImageHeaderFactory(blob);
    final effectiveFilename = filename ?? 'image.${header.defaultExt}';
    return Image._internal(blob, effectiveFilename, header);
 }

 /// Os bytes do 'arquivo' de imagem.
 Uint8List get blob => _blob;

 /// Tipo de conteúdo MIME para esta imagem, ex: 'image/jpeg'.
 String get contentType => _imageHeader.contentType;

 /// Extensão de arquivo para a imagem, sem o ponto inicial, ex: 'jpg'.
 String get ext {
    final parts = p.split(_filename);
    if (parts.isEmpty) return _imageHeader.defaultExt;
    final filenamePart = parts.last;
    final extIndex = filenamePart.lastIndexOf('.');
    return (extIndex != -1 && extIndex < filenamePart.length - 1)
        ? filenamePart.substring(extIndex + 1)
        : _imageHeader.defaultExt;
 }

 /// Nome do arquivo de imagem original ou um nome genérico.
 String get filename => _filename;

 /// Dimensão horizontal em pixels da imagem.
 int get pxWidth => _imageHeader.pxWidth;

 /// Dimensão vertical em pixels da imagem.
 int get pxHeight => _imageHeader.pxHeight;

 /// Pontos inteiros por polegada (DPI) para a largura desta imagem. Padrão 72.
 int get horzDpi => _imageHeader.horzDpi;

 /// Pontos inteiros por polegada (DPI) para a altura desta imagem. Padrão 72.
 int get vertDpi => _imageHeader.vertDpi;

 /// Valor [Length] representando a largura nativa da imagem.
 Length get width => Inches(pxWidth / horzDpi);

 /// Valor [Length] representando a altura nativa da imagem.
 Length get height => Inches(pxHeight / vertDpi);

 /// Par (cx, cy) representando as dimensões escaladas desta imagem.
 /// Retorna valores [Length].
 (Length, Length) scaledDimensions({Length? width, Length? height}) {
   // Implementação
    throw UnimplementedError();
 }

 /// Digest de hash SHA1 do blob da imagem.
 String get sha1 => crypto.sha1.convert(_blob).toString();
}

// Factory simulada
BaseImageHeader _ImageHeaderFactory(Uint8List blob) {
 // Lógica para detectar o tipo de imagem e chamar o construtor apropriado
 // Ex: if (isPng(blob)) return Png.fromBlob(blob);
 throw UnimplementedError("Image type detection not implemented");
}

// Definições para Bmp, Gif, etc. viriam aqui, herdando de BaseImageHeader
// e implementando os métodos/getters abstratos.
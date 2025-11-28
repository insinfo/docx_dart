// docx/image/image.dart
import 'dart:io';
import 'dart:typed_data';

import 'package:crypto/crypto.dart' as crypto; // Para sha1
import 'package:docx_dart/src/shared.dart'; // Para Length, Inches, Emu
import 'package:image/image.dart' as img; // Para metadados PNG
import 'package:path/path.dart' as p; // Para basename e splitext

const int _defaultPngDpi = 72;

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
    final nativeWidth = this.width;
    final nativeHeight = this.height;
    if (width == null && height == null) {
      return (nativeWidth, nativeHeight);
    }
    if (width != null && height != null) {
      return (width, height);
    }
    if (width != null) {
      final scale = nativeWidth.emu == 0 ? 1.0 : width.emu / nativeWidth.emu;
      final scaledHeight = nativeHeight * scale;
      return (width, scaledHeight);
    }
    final requestedHeight = height!;
    final scale =
        nativeHeight.emu == 0 ? 1.0 : requestedHeight.emu / nativeHeight.emu;
    final scaledWidth = nativeWidth * scale;
    return (scaledWidth, requestedHeight);
  }

  /// Digest de hash SHA1 do blob da imagem.
  String get sha1 => crypto.sha1.convert(_blob).toString();
}

// Factory simulada
BaseImageHeader _ImageHeaderFactory(Uint8List blob) {
  final pngDecoder = img.PngDecoder();
  if (pngDecoder.isValidFile(blob)) {
    final info = pngDecoder.startDecode(blob);
    if (info is! img.PngInfo) {
      throw FormatException('PNG stream missing header information');
    }
    return _PngImageHeader.fromInfo(info);
  }
  throw UnimplementedError('Only PNG images are supported at the moment');
}

class _PngImageHeader extends BaseImageHeader {
  _PngImageHeader(
      {required int pxWidth,
      required int pxHeight,
      required int horzDpi,
      required int vertDpi})
      : super(pxWidth, pxHeight, horzDpi, vertDpi);

  factory _PngImageHeader.fromInfo(img.PngInfo info) {
    final dims = info.pixelDimensions;
    return _PngImageHeader(
      pxWidth: info.width,
      pxHeight: info.height,
      horzDpi: _dpiFromPixelDimensions(dims, horizontal: true),
      vertDpi: _dpiFromPixelDimensions(dims, horizontal: false),
    );
  }

  @override
  String get contentType => 'image/png';

  @override
  String get defaultExt => 'png';
}

int _dpiFromPixelDimensions(img.PngPhysicalPixelDimensions? dims,
    {required bool horizontal}) {
  if (dims == null) {
    return _defaultPngDpi;
  }
  if (dims.unitSpecifier != img.PngPhysicalPixelDimensions.unitMeter) {
    return _defaultPngDpi;
  }
  final pxPerUnit = horizontal ? dims.xPxPerUnit : dims.yPxPerUnit;
  if (pxPerUnit <= 0) {
    return _defaultPngDpi;
  }
  final dpi = (pxPerUnit * 0.0254).round();
  return dpi > 0 ? dpi : _defaultPngDpi;
}

// Definições para Bmp, Gif, etc. viriam aqui, herdando de BaseImageHeader
// e implementando os métodos/getters abstratos.

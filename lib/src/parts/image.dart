import 'dart:typed_data';

import 'package:docx_dart/src/image/image.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/package.dart';
import 'package:docx_dart/src/opc/part.dart';

/// OPC image part wrapping raster data embedded in the package.
class ImagePart extends Part {
  Image? _image;

  ImagePart(
    PackUri partname,
    String contentType,
    Uint8List blob, {
    Image? image,
    OpcPackage? package,
  })  : _image = image,
        super(partname, contentType, blob, package);

  factory ImagePart.fromImage(Image image, PackUri partname) {
    return ImagePart(partname, image.contentType, image.blob, image: image);
  }

  static ImagePart load(
    PackUri partname,
    String contentType,
    Uint8List blob,
    OpcPackage package,
  ) {
    return ImagePart(partname, contentType, blob, package: package);
  }

  Image get image => _image ??= Image.fromBlob(blob);

  String get filename => _image?.filename ?? 'image.${partname.ext}';

  String get sha1 => image.sha1;
}

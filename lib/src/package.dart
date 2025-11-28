import 'dart:io';
import 'dart:typed_data';

import 'package:docx_dart/src/image/image.dart';
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/package.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/opc/pkgreader.dart';
import 'package:docx_dart/src/parts/image.dart';

class Package extends OpcPackage {
  ImageParts? _imageParts;

  static Package open(dynamic pkgFile) {
    final pkgReader = PackageReader.fromFile(pkgFile);
    final package = Package();
    Unmarshaller.unmarshal(pkgReader, package, PartFactory.newPart);
    return package;
  }

  @override
  void afterUnmarshal() {
    super.afterUnmarshal();
    _gatherImageParts();
  }

  ImagePart getOrAddImagePart(dynamic imageDescriptor) {
    return imageParts.getOrAddImagePart(imageDescriptor);
  }

  ImageParts get imageParts => _imageParts ??= ImageParts(this);

  void _gatherImageParts() {
    for (final rel in iterRels()) {
      if (rel.isExternal || rel.reltype != RELATIONSHIP_TYPE.IMAGE) {
        continue;
      }
      final targetPart = rel.targetPart;
      if (targetPart is! ImagePart) {
        continue;
      }
      if (imageParts.contains(targetPart)) {
        continue;
      }
      imageParts.append(targetPart);
    }
  }
}

class ImageParts {
  final Package _package;
  final List<ImagePart> _imageParts = [];

  ImageParts(this._package);

  bool contains(ImagePart part) => _imageParts.contains(part);

  void append(ImagePart part) => _imageParts.add(part);

  ImagePart getOrAddImagePart(dynamic descriptor) {
    final image = _imageFromDescriptor(descriptor);
    final existing = _getBySha1(image.sha1);
    if (existing != null) {
      return existing;
    }
    return _addImagePart(image);
  }

  ImagePart? _getBySha1(String sha1) {
    for (final imagePart in _imageParts) {
      if (imagePart.sha1 == sha1) {
        return imagePart;
      }
    }
    return null;
  }

  ImagePart _addImagePart(Image image) {
    final partname = _nextImagePartname(image.ext);
    final imagePart = ImagePart.fromImage(image, partname);
    append(imagePart);
    return imagePart;
  }

  PackUri _nextImagePartname(String ext) {
    final template = '/word/media/image%d.$ext';
    return _package.nextPartname(template);
  }

  Image _imageFromDescriptor(dynamic descriptor) {
    if (descriptor is Image) {
      return descriptor;
    }
    if (descriptor is Uint8List) {
      return Image.fromBytes(descriptor);
    }
    if (descriptor is List<int>) {
      return Image.fromBytes(Uint8List.fromList(descriptor));
    }
    if (descriptor is String) {
      final file = File(descriptor);
      final bytes = file.readAsBytesSync();
      return Image.fromBytes(Uint8List.fromList(bytes), filename: descriptor);
    }
    throw ArgumentError('Unsupported image descriptor type ${descriptor.runtimeType}');
  }
}

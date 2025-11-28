import 'dart:typed_data';

import 'package:docx_dart/src/default_docx_data.dart';
import 'package:docx_dart/src/document.dart' as docx;
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/oxml.dart' show parse_xml;
import 'package:docx_dart/src/opc/package.dart' as opc;
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/opc/parts/coreprops.dart';
import 'package:docx_dart/src/oxml/coreprops.dart';
import 'package:docx_dart/src/oxml/numbering.dart';
import 'package:docx_dart/src/package.dart';
import 'package:docx_dart/src/parts/document.dart';
import 'package:docx_dart/src/parts/hdrftr.dart';
import 'package:docx_dart/src/parts/numbering.dart';
import 'package:docx_dart/src/parts/settings.dart';
import 'package:docx_dart/src/parts/styles.dart';

/// Load a WordprocessingML document from [docxSource] or the built-in
/// default template when omitted.
///
/// [docxSource] can be a filesystem path, a `Uint8List`, a `List<int>`, or any
/// other value accepted by [Package.open].
///
/// The returned [docx.Document] can be manipulated and later saved using its
/// [docx.Document.save] method.
docx.Document loadDocxDocument([dynamic docxSource]) {
  _ensurePartLoadersRegistered();
  final resolvedSource = docxSource ?? defaultDocxBytes();
  final package = Package.open(resolvedSource);
  final mainPart = package.mainDocumentPart;
  if (mainPart is! DocumentPart) {
    throw StateError('Package main document part is not a DocumentPart.');
  }
  if (mainPart.contentType != CONTENT_TYPE.WML_DOCUMENT_MAIN) {
    throw ArgumentError(
      "Expected a WordprocessingML package, got content type '${mainPart.contentType}'.",
    );
  }
  return mainPart.document;
}

bool _partLoadersRegistered = false;

void _ensurePartLoadersRegistered() {
  if (_partLoadersRegistered) {
    return;
  }

  void register(String contentType, PartLoadFunction loader) {
    PartFactory.partTypeFor[contentType] = loader;
  }

  register(CONTENT_TYPE.WML_DOCUMENT_MAIN, _loadDocumentPart);
  register(CONTENT_TYPE.WML_STYLES, _loadStylesPart);
  register(CONTENT_TYPE.WML_SETTINGS, _loadSettingsPart);
  register(CONTENT_TYPE.WML_NUMBERING, _loadNumberingPart);
  register(CONTENT_TYPE.WML_HEADER, _loadHeaderPart);
  register(CONTENT_TYPE.WML_FOOTER, _loadFooterPart);
  register(CONTENT_TYPE.OPC_CORE_PROPERTIES, _loadCorePropertiesPart);

  _partLoadersRegistered = true;
}

Part _loadDocumentPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  return DocumentPart(partname, contentType, element, package);
}

Part _loadStylesPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  return StylesPart(partname, contentType, element, package);
}

Part _loadSettingsPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  return SettingsPart(partname, contentType, element, package);
}

Part _loadNumberingPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  final numberingElement =
      element is CT_Numbering ? element : CT_Numbering(element.element);
  return NumberingPart(partname, contentType, numberingElement, package);
}

Part _loadHeaderPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  return HeaderPart(partname, contentType, element, package);
}

Part _loadFooterPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  return FooterPart(partname, contentType, element, package);
}

Part _loadCorePropertiesPart(
  PackUri partname,
  String contentType,
  Uint8List blob,
  opc.OpcPackage package,
) {
  final element = parse_xml(blob);
  final coreElement =
      element is CT_CoreProperties ? element : CT_CoreProperties(element.element);
  return CorePropertiesPart(partname, contentType, coreElement, package);
}

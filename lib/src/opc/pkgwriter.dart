// docx/opc/pkgwriter.dart
import 'dart:typed_data';
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/oxml.dart' as opc_oxml;
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/phys_pkg.dart';
import 'package:docx_dart/src/opc/rel.dart';
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/opc/spec.dart'; // Para default_content_types
import 'package:docx_dart/src/opc/shared.dart'; // Para CaseInsensitiveMap
import 'package:docx_dart/src/oxml/xmlchemy.dart'; // Para serializePartXml

class PackageWriter {

 static void write(dynamic pkgFile, Relationships pkgRels, Iterable<Part> parts) {
    final physWriter = PhysPkgWriter(pkgFile);
    try {
      _writeContentTypesStream(physWriter, parts);
      _writePkgRels(physWriter, pkgRels);
      _writeParts(physWriter, parts);
    } finally {
       physWriter.close();
    }
 }

 static void _writeContentTypesStream(PhysPkgWriter physWriter, Iterable<Part> parts) {
    final cti = ContentTypesItem.fromParts(parts);
    physWriter.write(CONTENT_TYPES_URI, cti.blob);
 }

 static void _writeParts(PhysPkgWriter physWriter, Iterable<Part> parts) {
    for (final part in parts) {
      physWriter.write(part.partname, part.blob);
      if (part.rels.isNotEmpty) {
        physWriter.write(part.partname.relsUri, Uint8List.fromList(part.rels.xml.codeUnits)); // Assumindo que rels.xml retorna String
      }
    }
 }

 static void _writePkgRels(PhysPkgWriter physWriter, Relationships pkgRels) {
    physWriter.write(PACKAGE_URI.relsUri, Uint8List.fromList(pkgRels.xml.codeUnits)); // Assumindo que rels.xml retorna String
 }
}


class ContentTypesItem {
 final Map<String, String> _defaults = CaseInsensitiveMap<String>();
 final Map<String, String> _overrides = {}; // Partname é case-sensitive

 ContentTypesItem._();

 Uint8List get blob {
    final typesElm = _element;
    // Serializar typesElm para Uint8List usando package:xml
    return serializePartXml(typesElm);
 }

 factory ContentTypesItem.fromParts(Iterable<Part> parts) {
    final cti = ContentTypesItem._();
    cti._defaults['rels'] = CONTENT_TYPE.OPC_RELATIONSHIPS; // Exemplo, buscar de constantes
    cti._defaults['xml'] = CONTENT_TYPE.XML; // Exemplo

    for (final part in parts) {
      cti._addContentType(part.partname, part.contentType);
    }
    return cti;
 }

 void _addContentType(PackUri partname, String contentType) {
    final ext = partname.ext.toLowerCase();
    // Verifica se é um tipo padrão conhecido (usa opc.spec.default_content_types)
    final isDefault = defaultContentTypes.any((pair) => pair.$1 == ext && pair.$2 == contentType);

    if (isDefault) {
      _defaults[ext] = contentType;
    } else {
      _overrides[partname.uri] = contentType;
    }
 }

 opc_oxml.CT_Types get _element {
    final typesElm = opc_oxml.CT_Types.newTypes(); // Supondo um construtor estático
    final sortedDefaults = _defaults.entries.toList()..sort((a, b) => a.key.compareTo(b.key));
    final sortedOverrides = _overrides.entries.toList()..sort((a, b) => a.key.compareTo(b.key));

    for (final entry in sortedDefaults) {
      typesElm.add_default(entry.key, entry.value);
    }
    for (final entry in sortedOverrides) {
      typesElm.add_override(entry.key, entry.value);
    }
    return typesElm;
 }
}
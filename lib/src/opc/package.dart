import 'dart:typed_data';

import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/coreprops.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/opc/parts/coreprops.dart';
import 'package:docx_dart/src/opc/pkgreader.dart';
import 'package:docx_dart/src/opc/pkgwriter.dart';
import 'package:docx_dart/src/opc/rel.dart';

typedef PartBuilder = Part Function(
  PackUri partname,
  String contentType,
  String reltype,
  Uint8List blob,
  OpcPackage package,
);

class OpcPackage {
  Relationships? _rels;

  OpcPackage();

  void afterUnmarshal() {}

  CoreProperties get coreProperties => _corePropertiesPart.coreProperties;

  Iterable<Relationship> iterRels() sync* {
    Iterable<Relationship> walk(dynamic source, Set<Part> visited) sync* {
      final relationships = source.rels.values;
      for (final rel in relationships) {
        yield rel;
        if (rel.isExternal) {
          continue;
        }
        final part = rel.targetPart;
        if (visited.add(part)) {
          yield* walk(part, visited);
        }
      }
    }

    yield* walk(this, <Part>{});
  }

  Iterable<Part> iterParts() sync* {
    Iterable<Part> walk(dynamic source, Set<Part> visited) sync* {
      final relationships = source.rels.values;
      for (final rel in relationships) {
        if (rel.isExternal) {
          continue;
        }
        final part = rel.targetPart;
        if (!visited.add(part)) {
          continue;
        }
        yield part;
        yield* walk(part, visited);
      }
    }

    yield* walk(this, <Part>{});
  }

  Relationship loadRel(String reltype, Object target, String rId, {bool isExternal = false}) {
    return rels.addRelationship(reltype, target, rId, isExternal: isExternal);
  }

  Part get mainDocumentPart => partRelatedBy(RELATIONSHIP_TYPE.OFFICE_DOCUMENT);

  PackUri nextPartname(String template) {
    final partnames = {for (final part in iterParts()) part.partname.uri};
    for (var n = 1; n < partnames.length + 3; n++) {
      final candidate = template.replaceFirst('%d', '$n');
      if (!partnames.contains(candidate)) {
        return PackUri(candidate);
      }
    }
    // Fallback in the unlikely case template lacks %d or similar
    return PackUri(template);
  }

  static OpcPackage open(dynamic pkgFile) {
    final pkgReader = PackageReader.fromFile(pkgFile);
    final package = OpcPackage();
    Unmarshaller.unmarshal(pkgReader, package, PartFactory.newPart);
    return package;
  }

  Part partRelatedBy(String reltype) => rels.partWithReltype(reltype);

  List<Part> get parts => List<Part>.unmodifiable(iterParts());

  String relateTo(Part part, String reltype) => rels.getOrAdd(reltype, part).rId;

  Relationships get rels => _rels ??= Relationships(PACKAGE_URI.baseUri);

  void save(dynamic pkgFile) {
    for (final part in parts) {
      part.beforeMarshal();
    }
    PackageWriter.write(pkgFile, rels, parts);
  }

  CorePropertiesPart get _corePropertiesPart {
    try {
      return partRelatedBy(RELATIONSHIP_TYPE.CORE_PROPERTIES) as CorePropertiesPart;
    } catch (_) {
      final corePropertiesPart = CorePropertiesPart.defaultPart(this);
      relateTo(corePropertiesPart, RELATIONSHIP_TYPE.CORE_PROPERTIES);
      return corePropertiesPart;
    }
  }
}

class Unmarshaller {
  static void unmarshal(
    PackageReader pkgReader,
    OpcPackage package,
    PartBuilder partFactory,
  ) {
    final parts = _unmarshalParts(pkgReader, package, partFactory);
    _unmarshalRelationships(pkgReader, package, parts);
    for (final part in parts.values) {
      part.afterUnmarshal();
    }
    package.afterUnmarshal();
  }

  static Map<String, Part> _unmarshalParts(
    PackageReader pkgReader,
    OpcPackage package,
    PartBuilder partFactory,
  ) {
    final parts = <String, Part>{};
    for (final entry in pkgReader.iterSparts()) {
      final partname = entry.$1;
      final contentType = entry.$2;
      final reltype = entry.$3;
      final blob = entry.$4;
      parts[partname.uri] = partFactory(partname, contentType, reltype, blob, package);
    }
    return parts;
  }

  static void _unmarshalRelationships(
    PackageReader pkgReader,
    OpcPackage package,
    Map<String, Part> parts,
  ) {
    for (final entry in pkgReader.iterSrels()) {
      final sourceUri = entry.$1;
      final srel = entry.$2;
      final source = sourceUri.uri == PACKAGE_URI.uri ? package : parts[sourceUri.uri]!;
      final target = srel.isExternal ? srel.targetRef : parts[srel.targetPartname.uri]!;
      if (source is OpcPackage) {
        source.loadRel(srel.reltype, target, srel.rId, isExternal: srel.isExternal);
      } else if (source is Part) {
        source.loadRel(srel.reltype, target, srel.rId, isExternal: srel.isExternal);
      } else {
        throw StateError('Unsupported relationship source type: ${source.runtimeType}');
      }
    }
  }
}

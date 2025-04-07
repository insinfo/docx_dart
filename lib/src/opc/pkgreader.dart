// docx/opc/pkgreader.py
import 'dart:typed_data';
import 'package:xml/xml.dart';
import 'package:collection/collection.dart'; // Para firstWhereOrNull
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/oxml.dart' as opc_oxml; // Para parse_xml e classes CT_*
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/phys_pkg.dart';
import 'package:docx_dart/src/opc/shared.dart'; // Para CaseInsensitiveMap

class SerializedRelationship {
  final String _baseUri;
  final String rId;
  final String reltype;
  final RELATIONSHIP_TARGET_MODE targetMode;
  final String targetRef;
  PackUri? _targetPartname; // Cache

  SerializedRelationship._(this._baseUri, this.rId, this.reltype, this.targetMode, this.targetRef);

  factory SerializedRelationship.fromElement(String baseUri, opc_oxml.CT_Relationship relElm) {
    return SerializedRelationship._(
      baseUri,
      relElm.rId,
      relElm.reltype,
      RELATIONSHIP_TARGET_MODE.fromString(relElm.targetMode),
      relElm.targetRef,
    );
  }

  bool get isExternal => targetMode == RELATIONSHIP_TARGET_MODE.external;

  PackUri get targetPartname {
    if (isExternal) {
      throw StateError('targetPartname is undefined where TargetMode == External');
    }
    _targetPartname ??= PackUri.fromRelativeRef(_baseUri, targetRef);
    return _targetPartname!;
  }
}

class SerializedRelationships extends Iterable<SerializedRelationship> {
  final List<SerializedRelationship> _srels;

  SerializedRelationships._(this._srels);

  factory SerializedRelationships.loadFromXml(String baseUri, String? relsXml) {
    final srels = <SerializedRelationship>[];
    if (relsXml != null) {
      final relsElm = opc_oxml.parse_xml(relsXml) as opc_oxml.CT_Relationships;
      for (final relElm in relsElm.Relationship_lst) {
        srels.add(SerializedRelationship.fromElement(baseUri, relElm));
      }
    }
    return SerializedRelationships._(srels);
  }

  @override
  Iterator<SerializedRelationship> get iterator => _srels.iterator;
}


class ContentTypeMap {
 final Map<String, String> _overrides = CaseInsensitiveMap<String>();
 final Map<String, String> _defaults = CaseInsensitiveMap<String>();

 String operator [](PackUri partname) {
    final partnameStr = partname.uri;
    if (_overrides.containsKey(partnameStr)) {
      return _overrides[partnameStr]!;
    }
    final ext = partname.ext;
    if (_defaults.containsKey(ext)) {
      return _defaults[ext]!;
    }
    throw ArgumentError("no content type for partname '$partname' in [Content_Types].xml");
 }

 factory ContentTypeMap.fromXml(Uint8List contentTypesXml) {
    final typesElm = opc_oxml.parse_xml(String.fromCharCodes(contentTypesXml)) as opc_oxml.CT_Types;
    final ctMap = ContentTypeMap();
    for (final o in typesElm.overrides) {
      ctMap._addOverride(o.partname, o.content_type);
    }
    for (final d in typesElm.defaults) {
      ctMap._addDefault(d.extension, d.content_type);
    }
    return ctMap;
 }

 void _addDefault(String extension, String contentType) {
    _defaults[extension] = contentType;
 }

 void _addOverride(String partname, String contentType) {
    _overrides[partname] = contentType;
 }
}


class SerializedPart {
 final PackUri partname;
 final String contentType;
 final String reltype; // Relationship type *referring* to this part
 final Uint8List blob;
 final SerializedRelationships srels;

 SerializedPart(this.partname, this.contentType, this.reltype, this.blob, this.srels);
}


class PackageReader {
 final ContentTypeMap _contentTypes;
 final SerializedRelationships _pkgSrels;
 final List<SerializedPart> _sparts;

 PackageReader._(this._contentTypes, this._pkgSrels, this._sparts);

 static PackageReader fromFile(dynamic pkgFile) {
    final physReader = PhysPkgReader(pkgFile);
    try {
      final contentTypes = ContentTypeMap.fromXml(physReader.contentTypesXml);
      final pkgSrels = _srelsFor(physReader, PACKAGE_URI);
      final sparts = _loadSerializedParts(physReader, pkgSrels, contentTypes);
      return PackageReader._(contentTypes, pkgSrels, sparts);
    } finally {
      physReader.close();
    }
 }

 Iterable<(PackUri, String, String, Uint8List)> iterSparts() sync* {
    for (final s in _sparts) {
      yield (s.partname, s.contentType, s.reltype, s.blob);
    }
 }

 Iterable<(PackUri, SerializedRelationship)> iterSrels() sync* {
    for (final srel in _pkgSrels) {
      yield (PACKAGE_URI, srel);
    }
    for (final spart in _sparts) {
      for (final srel in spart.srels) {
        yield (spart.partname, srel);
      }
    }
 }

 static List<SerializedPart> _loadSerializedParts(
    PhysPkgReader physReader,
    SerializedRelationships pkgSrels,
    ContentTypeMap contentTypes,
  ) {
    final sparts = <SerializedPart>[];
    final visitedPartnames = <String>{};

    void walkParts(SerializedRelationships srels) {
      for (final srel in srels) {
        if (srel.isExternal) continue;
        final partname = srel.targetPartname;
        if (visitedPartnames.contains(partname.uri)) continue;

        visitedPartnames.add(partname.uri);
        final reltype = srel.reltype;
        final partSrels = _srelsFor(physReader, partname);
        final blob = physReader.blobFor(partname);
        final contentType = contentTypes[partname];

        final spart = SerializedPart(partname, contentType, reltype, blob, partSrels);
        sparts.add(spart);
        walkParts(partSrels); // Recurse
      }
    }

    walkParts(pkgSrels);
    return sparts;
 }


 static SerializedRelationships _srelsFor(PhysPkgReader physReader, PackUri sourceUri) {
    final relsXml = physReader.relsXmlFor(sourceUri);
    return SerializedRelationships.loadFromXml(sourceUri.baseUri, relsXml);
 }
}
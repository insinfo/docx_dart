import 'dart:typed_data';

import 'package:xml/xml.dart';

import 'package:docx_dart/src/opc/oxml.dart' show parse_xml, serializePartXml;
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/package.dart';
import 'package:docx_dart/src/opc/rel.dart';
import 'package:docx_dart/src/oxml/ns.dart' show nsmap;
import 'package:docx_dart/src/oxml/xmlchemy.dart';

typedef PartLoadFunction = Part Function(
  PackUri partname,
  String contentType,
  Uint8List blob,
  OpcPackage package,
);

typedef PartSelectorFunction = PartLoadFunction? Function(
  String contentType,
  String reltype,
);

class Part {
  PackUri _partname;
  final String _contentType;
  Uint8List? _blob;
  final OpcPackage? _package;
  Relationships? _rels;

  Part(this._partname, this._contentType, [this._blob, this._package]);

  void afterUnmarshal() {}

  void beforeMarshal() {}

  Uint8List get blob => _blob ?? Uint8List(0);

  String get contentType => _contentType;

  void dropRel(String rId) {
    if (_relRefCount(rId) < 2) {
      rels.remove(rId);
    }
  }

  static Part load(PackUri partname, String contentType, Uint8List blob, OpcPackage package) {
    return Part(partname, contentType, blob, package);
  }

  Relationship loadRel(String reltype, Object target, String rId, {bool isExternal = false}) {
    return rels.addRelationship(reltype, target, rId, isExternal: isExternal);
  }

  OpcPackage? get package => _package;

  PackUri get partname => _partname;

  set partname(PackUri value) => _partname = value;

  Part partRelatedBy(String reltype) => rels.partWithReltype(reltype);

  String relateTo(Object target, String reltype, {bool isExternal = false}) {
    if (isExternal) {
      if (target is! String) {
        throw ArgumentError('External relationships require a String target reference.');
      }
      return rels.getOrAddExternalRel(reltype, target);
    }
    if (target is! Part) {
      throw ArgumentError('Internal relationships require a Part target.');
    }
    return rels.getOrAdd(reltype, target).rId;
  }

  Map<String, Part> get relatedParts => rels.relatedParts;

  Relationships get rels => _rels ??= Relationships(_partname.baseUri);

  String targetRef(String rId) => rels[rId]!.targetRef;

  int _relRefCount(String rId) => 0;

  Part get part => this;
}

class PartFactory {
  static PartSelectorFunction? partClassSelector;
  static final Map<String, PartLoadFunction> partTypeFor = {};
  static PartLoadFunction defaultPartType = Part.load;

  static Part newPart(
    PackUri partname,
    String contentType,
    String reltype,
    Uint8List blob,
    OpcPackage package,
  ) {
    PartLoadFunction? loader;
    if (partClassSelector != null) {
      loader = partClassSelector!(contentType, reltype);
    }
    loader ??= partTypeFor[contentType];
    loader ??= defaultPartType;
    return loader(partname, contentType, blob, package);
  }
}

class XmlPart extends Part {
  final BaseOxmlElement _element;

  XmlPart(PackUri partname, String contentType, this._element, OpcPackage package)
      : super(partname, contentType, null, package);

  @override
  Uint8List get blob => serializePartXml(_element);

  BaseOxmlElement get element => _element;

  static XmlPart load(PackUri partname, String contentType, Uint8List blob, OpcPackage package) {
    final element = parse_xml(blob);
    return XmlPart(partname, contentType, element, package);
  }

  @override
  int _relRefCount(String rId) {
    final namespace = nsmap['r'];
    if (namespace == null) {
      return 0;
    }
    int count = 0;
    final Iterable<XmlElement> descendants = _element.element.descendants.whereType<XmlElement>();
    for (final descendant in descendants) {
      final attr = descendant.getAttribute('id', namespace: namespace);
      if (attr != null && attr == rId) {
        count += 1;
      }
    }
    return count;
  }
}

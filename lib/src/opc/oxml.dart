/// OPC-specific XML helpers and element wrappers.
///
/// Mirrors the limited functionality python-docx uses from docx.opc.oxml,
/// providing factories for content-types and relationships parts along with
/// convenience parsers/serializers shared across the OPC layer.

import 'dart:convert';
import 'dart:typed_data';

import 'package:xml/xml.dart';

import 'package:docx_dart/src/oxml/ns.dart' show nsmap, qn;
import 'package:docx_dart/src/oxml/xmlchemy.dart' show BaseOxmlElement;

/// Parses [source] (UTF-8 XML string or bytes) and wraps the root element in a
/// matching OPC helper class when possible. Falls back to [BaseOxmlElement]
/// when the element is not one of the specialized OPC types we recognize.
BaseOxmlElement parse_xml(Object source) {
  final String xmlString;
  if (source is Uint8List) {
    xmlString = utf8.decode(source);
  } else if (source is String) {
    xmlString = source;
  } else {
    throw ArgumentError('Unsupported XML source type: ${source.runtimeType}');
  }

  final document = XmlDocument.parse(xmlString);
  return _wrap(document.rootElement);
}

/// Serializes [element] (without pretty printing) into a UTF-8 encoded byte
/// buffer suitable for storage inside a ZIP part stream.
Uint8List serializePartXml(BaseOxmlElement element) => _serializeXmlElement(element.element);

/// Serializes a bare [XmlElement] to UTF-8 with the XML declaration.
Uint8List _serializeXmlElement(XmlElement element) {
  final body = element.toXmlString(pretty: false);
  final xml = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>$body";
  return Uint8List.fromList(utf8.encode(xml));
}

BaseOxmlElement _wrap(XmlElement element) {
  final uri = element.name.namespaceUri;
  final local = element.name.local;

  if (uri == nsmap['ct']) {
    switch (local) {
      case 'Types':
        return CT_Types(element);
      case 'Default':
        return CT_Default(element);
      case 'Override':
        return CT_Override(element);
    }
  }

  if (uri == nsmap['pr']) {
    switch (local) {
      case 'Relationships':
        return CT_Relationships(element);
      case 'Relationship':
        return CT_Relationship(element);
    }
  }

  return BaseOxmlElement(element);
}

class CT_Default extends BaseOxmlElement {
  CT_Default(XmlElement element) : super(element);

  factory CT_Default.newDefault(String ext, String contentType) {
    final xml = '<Default xmlns="${nsmap['ct']}"/>';
    final element = XmlDocument.parse(xml).rootElement;
    element.setAttribute('Extension', ext);
    element.setAttribute('ContentType', contentType);
    return CT_Default(element);
  }

  String get extension => element.getAttribute('Extension') ?? '';

  String get content_type => element.getAttribute('ContentType') ?? '';
}

class CT_Override extends BaseOxmlElement {
  CT_Override(XmlElement element) : super(element);

  factory CT_Override.newOverride(String partname, String contentType) {
    final xml = '<Override xmlns="${nsmap['ct']}"/>';
    final element = XmlDocument.parse(xml).rootElement;
    element.setAttribute('PartName', partname);
    element.setAttribute('ContentType', contentType);
    return CT_Override(element);
  }

  String get partname => element.getAttribute('PartName') ?? '';

  String get content_type => element.getAttribute('ContentType') ?? '';
}

class CT_Types extends BaseOxmlElement {
  CT_Types(XmlElement element) : super(element);

  factory CT_Types.newTypes() {
    final xml = '<Types xmlns="${nsmap['ct']}"/>';
    return CT_Types(XmlDocument.parse(xml).rootElement);
  }

  Iterable<CT_Default> get defaults => element
      .findElements('Default', namespace: nsmap['ct'])
      .map((node) => CT_Default(node));

  Iterable<CT_Override> get overrides => element
      .findElements('Override', namespace: nsmap['ct'])
      .map((node) => CT_Override(node));

  void add_default(String ext, String contentType) {
    element.children.add(CT_Default.newDefault(ext, contentType).element);
  }

  void add_override(String partname, String contentType) {
    element.children.add(CT_Override.newOverride(partname, contentType).element);
  }
}

class CT_Relationship extends BaseOxmlElement {
  CT_Relationship(XmlElement element) : super(element);

  factory CT_Relationship.newRelationship(String rId, String reltype, String target,
      {String targetMode = 'Internal'}) {
    final xml = '<Relationship xmlns="${nsmap['pr']}"/>';
    final element = XmlDocument.parse(xml).rootElement;
    element.setAttribute('Id', rId);
    element.setAttribute('Type', reltype);
    element.setAttribute('Target', target);
    if (targetMode != 'Internal') {
      element.setAttribute('TargetMode', targetMode);
    }
    return CT_Relationship(element);
  }

  String get rId => element.getAttribute('Id') ?? '';

  String get reltype => element.getAttribute('Type') ?? '';

  String get targetRef => element.getAttribute('Target') ?? '';

  String get targetMode => element.getAttribute('TargetMode') ?? 'Internal';
}

class CT_Relationships extends BaseOxmlElement {
  CT_Relationships(XmlElement element) : super(element);

  factory CT_Relationships.newRelationships() {
    final xml = '<Relationships xmlns="${nsmap['pr']}"/>';
    return CT_Relationships(XmlDocument.parse(xml).rootElement);
  }

  void add_rel(String rId, String reltype, String target, {bool isExternal = false}) {
    final targetMode = isExternal ? 'External' : 'Internal';
    final rel = CT_Relationship.newRelationship(rId, reltype, target, targetMode: targetMode);
    element.children.add(rel.element);
  }

  List<CT_Relationship> get Relationship_lst => element
      .findElements('Relationship', namespace: nsmap['pr'])
      .map((rel) => CT_Relationship(rel))
      .toList(growable: false);

  String get xml {
    final body = element.toXmlString(pretty: false);
    return "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>$body";
  }
}

/// Qualified name helper for OPC namespaces, identical to [qn] but kept here for
/// convenience by callers already importing this module.
String qn_opc(String tag) => qn(tag);

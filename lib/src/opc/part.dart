// docx/opc/part.dart
import 'dart:typed_data';
import 'package:docx_dart/src/opc/oxml.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/rel.dart';
import 'package:docx_dart/src/oxml/parser.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/opc/package.dart'; // Para OpcPackage
import 'package:docx_dart/src/opc/constants.dart'; // Para RELATIONSHIP_TARGET_MODE

typedef PartLoadFunction = Part Function(PackUri partname, String contentType, Uint8List blob, OpcPackage package);
typedef PartSelectorFunction = Type? Function(String contentType, String reltype); // Ou retorna PartLoadFunction?

class Part {
 PackUri _partname;
 final String _contentType;
 Uint8List? _blob; // Pode ser nulo se for XmlPart e ainda não serializado
 final OpcPackage? _package; // Pode ser nulo durante a construção inicial? Ajustar se necessário.
 late final Relationships _rels; // Usar late final

 Part(this._partname, this._contentType, [this._blob, this._package]) {
    _rels = Relationships(_partname.baseUri);
 }

 /// Chamado após unmarshalling para processamento adicional (ex: parse XML).
 void afterUnmarshal() {}

 /// Chamado antes da serialização para processamento (ex: finalizar nomes).
 void beforeMarshal() {}

 /// Conteúdo desta parte como bytes. Pode ser texto ou binário.
 Uint8List get blob => _blob ?? Uint8List(0); // Deve ser implementado por subclasses como XmlPart

 /// Tipo de conteúdo desta parte.
 String get contentType => _contentType;

 /// Remove o relacionamento identificado por `rId` se sua contagem de referências for menor que 2.
 void dropRel(String rId) {
   if (_relRefCount(rId) < 2) {
      _rels.remove(rId);
   }
 }

 /// Carrega uma parte do pacote. Usado por PartFactory.
 static Part load(PackUri partname, String contentType, Uint8List blob, OpcPackage package) {
    return Part(partname, contentType, blob, package);
 }

 /// Adiciona um novo relacionamento carregado do pacote.
 Relationship loadRel(String reltype, Object target, String rId, {bool isExternal = false}) {
    return rels.addRelationship(reltype, target, rId, isExternal: isExternal);
 }

 /// Instância [OpcPackage] à qual esta parte pertence.
 OpcPackage? get package => _package;

 /// Instância [PackUri] contendo o nome desta parte.
 PackUri get partname => _partname;
 set partname(PackUri value) {
    _partname = value;
    // Atualizar base URI dos relacionamentos se necessário?
    _rels = Relationships(_partname.baseUri); // Recriar com nova base URI
 }

 /// Retorna a parte relacionada por `reltype`. Lança exceção se não encontrada ou múltipla.
 Part partRelatedBy(String reltype) {
    return rels.partWithReltype(reltype);
 }

 /// Adiciona ou obtém relacionamento com `target`. Retorna rId.
 String relateTo(Part target, String reltype, {bool isExternal = false}) {
    if (isExternal) {
       return rels.getOrAddExternalRel(reltype, target.partname.uri); // Assumindo que target é string URL se external
    } else {
       return rels.getOrAdd(reltype, target).rId;
    }
 }

 /// Mapa de rId para partes relacionadas (alvos internos).
 Map<String, Part> get relatedParts => rels.relatedParts;

 /// Coleção de relacionamentos para esta parte.
 Relationships get rels => _rels;

 /// Retorna a referência de destino (URL ou path relativo) para `rId`.
 String targetRef(String rId) => rels[rId]!.targetRef;

 /// Contagem de referências a `rId` dentro desta parte (0 para não-XmlPart).
 int _relRefCount(String rId) => 0;
}


class PartFactory {
 static PartSelectorFunction? partClassSelector;
 static Map<String, PartLoadFunction> partTypeFor = {};
 static PartLoadFunction defaultPartType = Part.load;

 static Part newPart(PackUri partname, String contentType, String reltype, Uint8List blob, OpcPackage package) {
    PartLoadFunction? loader;

    if (partClassSelector != null) {
       final selectedType = partClassSelector!(contentType, reltype);
       // Lógica para mapear Type para PartLoadFunction se necessário,
       // ou fazer partClassSelector retornar PartLoadFunction diretamente.
       // Por simplicidade, vamos assumir que a lógica de mapeamento está implícita.
       if (selectedType != null) {
          // loader = findLoaderForType(selectedType); // Lógica hipotética
          print("Warning: PartSelectorFunction logic needs refinement in Dart");
       }
    }

    if (loader == null && partTypeFor.containsKey(contentType)) {
      loader = partTypeFor[contentType];
    }

    loader ??= defaultPartType;

    return loader(partname, contentType, blob, package);
 }
}


class XmlPart extends Part {
 BaseOxmlElement _element; // O elemento raiz XML

 XmlPart(PackUri partname, String contentType, this._element, OpcPackage package)
    : super(partname, contentType, null, package); // Blob inicial é nulo

 @override
 Uint8List get blob => serializePartXml(_element); // Serializa o elemento XML

 BaseOxmlElement get element => _element;

 /// Carrega uma XmlPart a partir de um blob XML.
 static XmlPart load(PackUri partname, String contentType, Uint8List blob, OpcPackage package) {
   final element = parse_xml(blob); // Usa o parser OX comunitàrio
   // TODO: Precisa de um mecanismo para determinar a classe XmlPart específica (ex: DocumentPart)
   // Isso pode ser feito em PartFactory ou aqui com base no contentType.
   // Por agora, retornamos a classe base.
   return XmlPart(partname, contentType, element, package);
 }

 @override
 XmlPart get part => this; // Terminus para delegação de part

 @override
 int _relRefCount(String rId) {
   // Implementação usando xpath no _element
   final rIds = _element.xpath("//@r:id").map((attr) => attr.value).toList();
   return rIds.where((_rId) => _rId == rId).length;
 }
}

// Funções utilitárias simuladas para serialização/parsing
Uint8List serializePartXml(BaseOxmlElement element) {
  // Implementação usando package:xml para serializar
  throw UnimplementedError();
}
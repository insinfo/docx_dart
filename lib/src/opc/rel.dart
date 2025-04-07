// docx/opc/rel.dart
import 'dart:collection'; // Para HashMap
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/oxml.dart'; // Para CT_Relationships
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/opc/packuri.dart';

class Relationship {
 final String _rId;
 final String _reltype;
 final Object _target; // Pode ser Part ou String (URL)
 final String _baseUri;
 final bool _isExternal;

 Relationship(this._rId, this._reltype, this._target, this._baseUri, [this._isExternal = false]);

 bool get isExternal => _isExternal;
 String get reltype => _reltype;
 String get rId => _rId;

 Part get targetPart {
    if (_isExternal) {
      throw StateError("target_part is undefined when target mode is External");
    }
    return _target as Part;
 }

 String get targetRef {
    if (_isExternal) {
      return _target as String;
    } else {
      return (_target as Part).partname.relativeRef(_baseUri);
    }
 }
}


class Relationships extends MapBase<String, Relationship> {
  final String _baseUri;
  final Map<String, Relationship> _rels = {};
  // _target_parts_by_rId não é diretamente necessário se Map for usado

  Relationships(this._baseUri);

  @override
  Relationship? operator [](Object? key) => _rels[key];

  @override
  void operator []=(String key, Relationship value) => _rels[key] = value;

  @override
  void clear() => _rels.clear();

  @override
  Iterable<String> get keys => _rels.keys;

  @override
  Relationship? remove(Object? key) => _rels.remove(key);

  Relationship addRelationship(String reltype, Object target, String rId, {bool isExternal = false}) {
     final rel = Relationship(rId, reltype, target, _baseUri, isExternal);
     this[rId] = rel;
     return rel;
  }

  Relationship getOrAdd(String reltype, Part targetPart) {
     final existing = _getMatching(reltype, targetPart, isExternal: false);
     if (existing != null) return existing;
     final rId = _nextRId;
     return addRelationship(reltype, targetPart, rId, isExternal: false);
  }

   String getOrAddExternalRel(String reltype, String targetRef) {
      final existing = _getMatching(reltype, targetRef, isExternal: true);
      if (existing != null) return existing.rId;
      final rId = _nextRId;
      return addRelationship(reltype, targetRef, rId, isExternal: true).rId;
   }

   Part partWithReltype(String reltype) {
     final rel = _getRelOfType(reltype);
     return rel.targetPart;
   }

   /// Retorna partes relacionadas internas.
   Map<String, Part> get relatedParts {
      final parts = <String, Part>{};
      for (final rel in values) {
         if (!rel.isExternal) {
            parts[rel.rId] = rel.targetPart;
         }
      }
      return parts;
   }

   String get xml {
     // Implementação usando CT_Relationships de oxml.dart
     throw UnimplementedError();
   }

  Relationship? _getMatching(String reltype, Object target, {bool isExternal = false}) {
     for (final rel in values) {
        if (rel.reltype == reltype && rel.isExternal == isExternal) {
           final relTarget = rel.isExternal ? rel.targetRef : rel.targetPart;
           if (relTarget == target) {
              return rel;
           }
        }
     }
     return null;
  }

   Relationship _getRelOfType(String reltype) {
      final matching = values.where((rel) => rel.reltype == reltype).toList();
      if (matching.isEmpty) {
        throw ArgumentError("no relationship of type '$reltype' in collection");
      }
      if (matching.length > 1) {
         throw StateError("multiple relationships of type '$reltype' in collection");
      }
      return matching.first;
   }

   String get _nextRId {
      int n = 1;
      while (true) {
         final rIdCandidate = 'rId$n';
         if (!containsKey(rIdCandidate)) {
            return rIdCandidate;
         }
         n++;
      }
   }
}
/// Path: lib/src/opc/shared.dart
/// Based on python-docx: docx/opc/shared.py
///
/// Objects shared by OPC modules.

import 'dart:collection'; // Required for MapBase and HashMap

/// A Map implementation based on MapBase that treats String keys
/// case-insensitively.
///
/// For example, `cim['A']` and `cim['a']` access the same entry.
/// It uses an internal HashMap for storage.
class CaseInsensitiveMap<V> extends MapBase<String, V> {
  final Map<String, V> _innerMap = HashMap<String, V>();

  /// Creates an empty case-insensitive map.
  CaseInsensitiveMap();

  /// Creates a case-insensitive map that contains all key/value pairs of [other].
  CaseInsensitiveMap.from(Map<String, V> other) {
    addAll(other); // Use the addAll which uses our overridden []=
  }

  // --- Core methods required by MapBase ---

  @override
  V? operator [](Object? key) {
    if (key is String) {
      return _innerMap[key.toLowerCase()];
    }
    return null; // Only support String keys for case-insensitivity
  }

  @override
  void operator []=(String key, V value) {
    _innerMap[key.toLowerCase()] = value;
  }

  @override
  void clear() {
    _innerMap.clear();
  }

  @override
  Iterable<String> get keys => _innerMap.keys;
  // Note: The keys returned will be the lowercase versions used internally.
  // If you need the original case preserved (e.g., from the last insertion),
  // you'd need a more complex internal structure, perhaps another map storing
  // original key cases mapped to lowercase keys. For OPC needs, lowercase
  // keys are likely sufficient.

  @override
  V? remove(Object? key) {
    if (key is String) {
      return _innerMap.remove(key.toLowerCase());
    }
    return null; // Only support String keys for removal
  }

  // --- Optional overrides for efficiency or specific behavior ---

  // containsKey is provided by MapBase based on [] lookup,
  // but overriding can be slightly more direct/efficient.
  @override
  bool containsKey(Object? key) {
    if (key is String) {
      return _innerMap.containsKey(key.toLowerCase());
    }
    return false;
  }

  // length is provided by MapBase based on keys.length
  @override
  int get length => _innerMap.length;

  // isEmpty is provided by MapBase based on length
  @override
  bool get isEmpty => _innerMap.isEmpty;

  // isNotEmpty is provided by MapBase based on length
  @override
  bool get isNotEmpty => _innerMap.isNotEmpty;

}

// Note: The Python function `cls_method_fn` uses dynamic attribute access
// (`getattr`) which doesn't have a direct, commonly used, non-reflection
// equivalent in Dart. Dart typically relies on static dispatch or passing
// function references directly. Therefore, `cls_method_fn` is not directly
// translated here. The functionality it provides (likely dynamic method
// lookup for PartFactory) would need to be implemented differently in the
// Dart version of PartFactory, probably using a Map or similar registry pattern.
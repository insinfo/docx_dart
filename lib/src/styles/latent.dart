// docx/styles/latent.dart
// Partial Dart port of python-docx/docx/styles/latent.py

import 'dart:collection';

import 'package:collection/collection.dart';
import 'package:xml/xml.dart';

import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/styles/babel_fish.dart';
import 'package:docx_dart/src/types.dart';

class LatentStyles extends ElementProxy with IterableMixin<_LatentStyle> {
  LatentStyles(BaseOxmlElement element, [ProvidesXmlPart? parent])
      : super(element, parent);

  @override
  Iterator<_LatentStyle> get iterator => _lsdElements
      .map((lsd) => _LatentStyle(BaseOxmlElement(lsd), this))
      .iterator;

  @override
  int get length => _lsdElements.length;

  _LatentStyle operator [](String name) {
    final internal = BabelFish.ui2internal(name);
    final match = _lsdElements.firstWhereOrNull(
      (node) => node.getAttribute('name', namespace: nsmap['w']) == internal,
    );
    if (match == null) {
      throw ArgumentError("No latent style with name '$name'.");
    }
    return _LatentStyle(BaseOxmlElement(match), this);
  }

  _LatentStyle addLatentStyle(String name) {
    final internal = BabelFish.ui2internal(name);
    final newElement = XmlElement(XmlName('lsdException', 'w'));
    newElement.setAttribute('name', internal, namespace: nsmap['w']);
    element.element.children.add(newElement);
    return _LatentStyle(BaseOxmlElement(newElement), this);
  }

  int? get defaultPriority => _getIntAttr('defUIPriority');
  set defaultPriority(int? value) => _setIntAttr('defUIPriority', value);

  bool get defaultToHidden => _getBoolAttr('defSemiHidden') ?? false;
  set defaultToHidden(bool value) =>
      _setBoolAttr('defSemiHidden', value, removeOnFalse: true);

  bool get defaultToLocked => _getBoolAttr('defLockedState') ?? false;
  set defaultToLocked(bool value) =>
      _setBoolAttr('defLockedState', value, removeOnFalse: true);

  bool get defaultToQuickStyle => _getBoolAttr('defQFormat') ?? false;
  set defaultToQuickStyle(bool value) =>
      _setBoolAttr('defQFormat', value, removeOnFalse: true);

  bool get defaultToUnhideWhenUsed => _getBoolAttr('defUnhideWhenUsed') ?? false;
  set defaultToUnhideWhenUsed(bool value) =>
      _setBoolAttr('defUnhideWhenUsed', value, removeOnFalse: true);

  int? get loadCount => _getIntAttr('count');
  set loadCount(int? value) => _setIntAttr('count', value);

  List<XmlElement> get _lsdElements => element.element
      .findElements('lsdException', namespace: nsmap['w'])
      .toList(growable: false);

  int? _getIntAttr(String name) {
    final raw = element.element.getAttribute(name, namespace: nsmap['w']);
    return raw == null ? null : int.tryParse(raw);
  }

  void _setIntAttr(String name, int? value) {
    if (value == null) {
      element.element.removeAttribute(name, namespace: nsmap['w']);
    } else {
      element.element.setAttribute(name, value.toString(), namespace: nsmap['w']);
    }
  }

  bool? _getBoolAttr(String name) {
    final raw = element.element.getAttribute(name, namespace: nsmap['w']);
    if (raw == null) {
      return null;
    }
    final normalized = raw.toLowerCase();
    if (normalized == '1' || normalized == 'true') {
      return true;
    }
    if (normalized == '0' || normalized == 'false') {
      return false;
    }
    return null;
  }

  void _setBoolAttr(String name, bool value, {bool removeOnFalse = false}) {
    if (!value && removeOnFalse) {
      element.element.removeAttribute(name, namespace: nsmap['w']);
      return;
    }
    element.element
        .setAttribute(name, value ? '1' : '0', namespace: nsmap['w']);
  }
}

class _LatentStyle extends ElementProxy {
  _LatentStyle(BaseOxmlElement element, [ProvidesXmlPart? parent])
      : super(element, parent);

  void delete() {
    element.element.parent?.children.remove(element.element);
  }

  bool? get hidden => _getTriState('semiHidden');
  set hidden(bool? value) => _setTriState('semiHidden', value);

  bool? get locked => _getTriState('locked');
  set locked(bool? value) => _setTriState('locked', value);

  String get name =>
      BabelFish.internal2ui(element.element.getAttribute('name', namespace: nsmap['w']) ?? '');

  int? get priority => _getIntAttr('uiPriority');
  set priority(int? value) => _setIntAttr('uiPriority', value);

  bool? get quickStyle => _getTriState('qFormat');
  set quickStyle(bool? value) => _setTriState('qFormat', value);

  bool? get unhideWhenUsed => _getTriState('unhideWhenUsed');
  set unhideWhenUsed(bool? value) => _setTriState('unhideWhenUsed', value);

  bool? _getTriState(String name) {
    final raw = element.element.getAttribute(name, namespace: nsmap['w']);
    if (raw == null) {
      return null;
    }
    final normalized = raw.toLowerCase();
    if (normalized == '1' || normalized == 'true') {
      return true;
    }
    if (normalized == '0' || normalized == 'false') {
      return false;
    }
    return null;
  }

  void _setTriState(String name, bool? value) {
    if (value == null) {
      element.element.removeAttribute(name, namespace: nsmap['w']);
      return;
    }
    element.element
        .setAttribute(name, value ? '1' : '0', namespace: nsmap['w']);
  }

  int? _getIntAttr(String name) {
    final raw = element.element.getAttribute(name, namespace: nsmap['w']);
    return raw == null ? null : int.tryParse(raw);
  }

  void _setIntAttr(String name, int? value) {
    if (value == null) {
      element.element.removeAttribute(name, namespace: nsmap['w']);
    } else {
      element.element.setAttribute(name, value.toString(), namespace: nsmap['w']);
    }
  }
}

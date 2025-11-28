// styles/styles.dart
// Dart port of python-docx/docx/styles/styles.py

import 'dart:collection';

import 'package:collection/collection.dart';
import 'package:xml/xml.dart';

import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/styles/babel_fish.dart';
import 'package:docx_dart/src/styles/latent.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/types.dart';

/// Provides access to the styles defined in a document's styles part.
///
/// Accessed using the `Document.styles` property (or `StylesPart.styles`).
/// Supports `length`, iteration, and dictionary-style access by style name.
class Styles extends ElementProxy with IterableMixin<BaseStyle> {
  LatentStyles? _latentStyles;

  Styles(BaseOxmlElement element, [ProvidesXmlPart? parent])
      : super(element, parent);

  @override
  Iterator<BaseStyle> get iterator => _styleElements
      .map((el) => StyleFactory.create(BaseOxmlElement(el), this))
      .iterator;

  @override
  int get length => _styleElements.length;

  @override
  bool contains(Object? element) {
    if (element is! String) {
      return false;
    }
    return containsName(element);
  }

  bool containsName(String name) {
    final internal = BabelFish.ui2internal(name);
    return _styleElements.any((el) => _styleName(el) == internal);
  }

  BaseStyle operator [](String key) {
    final internal = BabelFish.ui2internal(key);
    final byName =
        _styleElements.firstWhereOrNull((el) => _styleName(el) == internal);
    if (byName != null) {
      return StyleFactory.create(BaseOxmlElement(byName), this);
    }

    final byId = _styleElements
        .firstWhereOrNull((el) => _styleId(el)?.toLowerCase() == key.toLowerCase());
    if (byId != null) {
      return StyleFactory.create(BaseOxmlElement(byId), this);
    }

    throw ArgumentError("No style with name or styleId '$key'.");
  }

  BaseStyle addStyle(String name, WD_STYLE_TYPE type, {bool builtin = false}) {
    if (containsName(name)) {
      throw ArgumentError("Document already contains style '$name'.");
    }
    final internal = BabelFish.ui2internal(name);
    final styleId = StyleFactory.styleIdFromStyleName(internal);
    final styleEl = XmlElement(
      XmlName('style', 'w'),
      [
        XmlAttribute(XmlName('type', 'w'), type.xmlValue),
        XmlAttribute(XmlName('styleId', 'w'), styleId),
        if (!builtin)
          XmlAttribute(XmlName('customStyle', 'w'), '1'),
      ],
      [
        XmlElement(
          XmlName('name', 'w'),
          [XmlAttribute(XmlName('val', 'w'), internal)],
        ),
      ],
    );
    element.element.children.add(styleEl);
    return StyleFactory.create(BaseOxmlElement(styleEl), this);
  }

  BaseStyle? defaultStyleFor(WD_STYLE_TYPE type) {
    final match = _styleElements.lastWhereOrNull(
      (el) => _styleType(el) == type && _isDefault(el),
    );
    return match == null
        ? null
        : StyleFactory.create(BaseOxmlElement(match), this);
  }

  BaseStyle? getById(String? styleId, WD_STYLE_TYPE styleType) {
    if (styleId == null) {
      return defaultStyleFor(styleType);
    }
    final match = _styleElements.firstWhereOrNull(
      (el) => _styleId(el)?.toLowerCase() == styleId.toLowerCase(),
    );
    if (match == null || _styleType(match) != styleType) {
      return defaultStyleFor(styleType);
    }
    return StyleFactory.create(BaseOxmlElement(match), this);
  }

  String? getStyleId(dynamic styleOrName, WD_STYLE_TYPE styleType) {
    if (styleOrName == null) {
      return null;
    }
    if (styleOrName is BaseStyle) {
      if (styleOrName.type != styleType) {
        throw ArgumentError(
          'Assigned style is type ${styleOrName.type}, need type $styleType',
        );
      }
      final defaultId = defaultStyleFor(styleType)?.styleId;
      if (styleOrName.styleId == defaultId) {
        return null;
      }
      return styleOrName.styleId;
    }
    if (styleOrName is String) {
      final resolved = this[styleOrName];
      return getStyleId(resolved, styleType);
    }
    throw ArgumentError(
        'Expected style name or BaseStyle, got ${styleOrName.runtimeType}.');
  }

  LatentStyles get latentStyles {
    if (_latentStyles != null) {
      return _latentStyles!;
    }
    var node = element.element
        .findElements('latentStyles', namespace: nsmap['w'])
        .firstOrNull;
    if (node == null) {
      node = XmlElement(XmlName('latentStyles', 'w'));
      final docDefaults = element.element
          .findElements('docDefaults', namespace: nsmap['w'])
          .firstOrNull;
      if (docDefaults != null) {
        final index = element.element.children.indexOf(docDefaults);
        element.element.children.insert(index + 1, node);
      } else {
        element.element.children.insert(0, node);
      }
    }
    _latentStyles = LatentStyles(BaseOxmlElement(node), this);
    return _latentStyles!;
  }

  List<XmlElement> get _styleElements => element.element
      .findElements('style', namespace: nsmap['w'])
      .toList(growable: false);

  String? _styleName(XmlElement el) => el
      .findElements('name', namespace: nsmap['w'])
      .firstOrNull
      ?.getAttribute('val', namespace: nsmap['w']);

  String? _styleId(XmlElement el) =>
      el.getAttribute('styleId', namespace: nsmap['w']);

  WD_STYLE_TYPE _styleType(XmlElement el) {
    final raw = el.getAttribute('type', namespace: nsmap['w']);
    return raw == null ? WD_STYLE_TYPE.PARAGRAPH : WD_STYLE_TYPE.fromXml(raw);
  }

  bool _isDefault(XmlElement el) {
    final attr = el.getAttribute('default', namespace: nsmap['w']);
    return attr == '1' || attr?.toLowerCase() == 'true';
  }
}
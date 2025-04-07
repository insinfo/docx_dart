// styles/styles.dart
// Corresponds to Python file: python-docx/src/docx/styles/styles.py

import 'dart:collection';
import 'package:docx_dart/src/enum/style.dart';
import 'package:xml/xml.dart';
import 'package:collection/collection.dart'; // For firstWhereOrNull, etc.

// Assuming ElementProxy is defined, if not, wrap XmlElement directly
import '../shared.dart'; // For ElementProxy (or define base wrapper)


import '../oxml/ns.dart'; // For qn function and ns map
import 'babel_fish.dart'; // For name translation
import 'latent.dart'; // For LatentStyles
import 'style.dart'; // For BaseStyle and StyleFactory

/// Provides access to the styles defined in a document's styles part.
///
/// Accessed using the `Document.styles` property (or `StylesPart.styles`).
/// Supports `length`, iteration, and dictionary-style access by style name.
class Styles extends ElementProxy with IterableMixin<BaseStyle> {
  late final XmlElement _stylesElement;
  LatentStyles? _latentStyles; // Cache for lazy loading

  Styles(XmlElement stylesElement) : super(stylesElement) {
    _stylesElement = stylesElement;
  }

  /// Enables `styles.contains('Style Name')` check.
  bool contains(String name) {
    final internalName = BabelFish.ui2internal(name);
    return _styleElementList
        .any((el) => _getStyleName(el) == internalName);
  }

  /// Enables dictionary-style access by UI style name (e.g., `styles['Heading 1']`).
  ///
  /// Lookup by style ID is deprecated and will print a warning.
  @override
  BaseStyle operator [](String key) {
    // 1. Try lookup by name
    final internalName = BabelFish.ui2internal(key);
    var styleElm = _styleElementList
        .firstWhereOrNull((el) => _getStyleName(el) == internalName);

    if (styleElm != null) {
      return StyleFactory.create(styleElm);
    }

    // 2. Try lookup by styleId (deprecated)
    styleElm = _getElementById(key);
    if (styleElm != null) {
      print( // Using print for warning, consider a logger in production
        '[WARNING] Style lookup by style_id "$key" is deprecated. Use style name instead.'
      );
      return StyleFactory.create(styleElm);
    }

    throw ArgumentError("no style with name or styleId '$key'");
  }

  /// Provides iteration over the styles in document order (e.g., `for (var style in styles)`).
  @override
  Iterator<BaseStyle> get iterator =>
      _styleElementList.map((el) => StyleFactory.create(el)).iterator;

  /// The number of styles defined in the document.
  @override
  int get length => _styleElementList.length;

  /// Returns a newly added [BaseStyle] object of [styleType] and identified by [name].
  ///
  /// A builtin style can be indicated by passing `true` for the optional [builtin]
  /// argument. Throws [ArgumentError] if a style with the same name already exists.
  BaseStyle addStyle(String name, WD_STYLE_TYPE styleType, {bool builtin = false}) {
    final uiName = name; // Keep original for error message
    final internalName = BabelFish.ui2internal(name);

    if (contains(uiName)) { // Check using UI name for user clarity
      throw ArgumentError("Document already contains style '$uiName'");
    }

    final styleId = StyleFactory.styleIdFromStyleName(internalName);

    // Create the new <w:style> element
    final styleEl = XmlElement(
      XmlName('style', namespace: ns['w']),
      [
        XmlAttribute(XmlName('type', namespace: ns['w']), styleType.xmlValue),
        XmlAttribute(XmlName('styleId', namespace: ns['w']), styleId),
        if (!builtin) // Add customStyle attribute only if *not* builtin
          XmlAttribute(XmlName('customStyle', namespace: ns['w']), '1'),
      ],
      [
        XmlElement( // Add <w:name> child
          XmlName('name', namespace: ns['w']),
          [XmlAttribute(XmlName('val', namespace: ns['w']), internalName)],
          [],
        ),
      ],
    );

    // Insert into <w:styles> (usually just append)
    _stylesElement.children.add(styleEl);

    return StyleFactory.create(styleEl);
  }

  /// Returns the default style for [styleType] or `null` if no default is
  /// defined for that type (uncommon).
  BaseStyle? defaultStyleFor(WD_STYLE_TYPE styleType) {
    final defaultEl = _styleElementList.lastWhereOrNull(
      (el) => _getStyleType(el) == styleType && _isDefaultStyle(el)
    );

    return defaultEl == null ? null : StyleFactory.create(defaultEl);
  }

  /// Returns the style of [styleType] matching [styleId].
  ///
  /// Returns the default style for [styleType] if [styleId] is `null`,
  /// the style is not found, or if the found style is not of [styleType].
  BaseStyle? getById(String? styleId, WD_STYLE_TYPE styleType) {
    if (styleId == null) {
      return defaultStyleFor(styleType);
    }
    return _getById(styleId, styleType);
  }

  /// Returns the style ID (String) for the given [styleOrName].
  ///
  /// Returns `null` if [styleOrName] is `null` or resolves to the default style.
  /// [styleOrName] can be a style name (String) or a [BaseStyle] object.
  /// Throws [ArgumentError] if the style is not found or is of the wrong type.
  String? getStyleId(dynamic styleOrName /* String? | BaseStyle? */, WD_STYLE_TYPE styleType) {
    if (styleOrName == null) {
      return null;
    } else if (styleOrName is BaseStyle) {
      return _getStyleIdFromStyle(styleOrName, styleType);
    } else if (styleOrName is String) {
      return _getStyleIdFromName(styleOrName, styleType);
    } else {
      throw ArgumentError('Expected String or BaseStyle, got ${styleOrName.runtimeType}');
    }
  }

  /// A [LatentStyles] object providing access to latent style defaults and exceptions.
  LatentStyles get latentStyles {
    if (_latentStyles == null) {
      var latentStylesEl = _stylesElement.findElements('latentStyles', namespace: ns['w']).firstOrNull;
      if (latentStylesEl == null) {
        // Create and add <w:latentStyles> if it doesn't exist
        latentStylesEl = XmlElement(XmlName('latentStyles', namespace: ns['w']));
        // Insert according to schema order (after docDefaults, before style)
        final docDefaults = _stylesElement.findElements('docDefaults', namespace: ns['w']).firstOrNull;
        if (docDefaults != null) {
            final index = _stylesElement.children.indexOf(docDefaults);
             _stylesElement.children.insert(index + 1, latentStylesEl);
        } else {
            _stylesElement.children.insert(0, latentStylesEl); // Or append if schema allows
        }
      }
      _latentStyles = LatentStyles(latentStylesEl);
    }
    return _latentStyles!;
  }

  // --- Private Helpers ---

  /// Internal lookup by ID, returning the default style if not found or wrong type.
  BaseStyle? _getById(String styleId, WD_STYLE_TYPE styleType) {
    final styleEl = _getElementById(styleId);
    if (styleEl == null || _getStyleType(styleEl) != styleType) {
      return defaultStyleFor(styleType);
    }
    return StyleFactory.create(styleEl);
  }

   /// Internal lookup by Name, returning the default style if not found or wrong type.
  BaseStyle? _getByName(String internalName, WD_STYLE_TYPE styleType) {
    final styleEl = _getElementByName(internalName);
    if (styleEl == null || _getStyleType(styleEl) != styleType) {
      return defaultStyleFor(styleType);
    }
    return StyleFactory.create(styleEl);
  }


  /// Resolves style ID from a style name.
  String? _getStyleIdFromName(String styleName, WD_STYLE_TYPE styleType) {
    final style = this[styleName]; // Use operator[] which handles lookup & exceptions
    return _getStyleIdFromStyle(style, styleType);
  }

  /// Resolves style ID from a style object.
  String? _getStyleIdFromStyle(BaseStyle style, WD_STYLE_TYPE styleType) {
    if (style.type != styleType) {
      throw ArgumentError(
          "Assigned style is type ${style.type}, need type $styleType");
    }
    // Check if the provided style is the default for its type
    if (style.styleId == defaultStyleFor(styleType)?.styleId) {
      return null;
    }
    return style.styleId;
  }


  /// Gets the underlying list of <w:style> elements.
  List<XmlElement> get _styleElementList =>
      _stylesElement.findElements('style', namespace: ns['w']).toList();

  /// Finds a <w:style> element by its w:styleId attribute. Returns null if not found.
  XmlElement? _getElementById(String styleId) {
     return _styleElementList.firstWhereOrNull(
        (el) => el.getAttribute('styleId', namespace: ns['w']) == styleId);
  }

   /// Finds a <w:style> element by its <w:name w:val="..."/> grandchild. Returns null if not found.
  XmlElement? _getElementByName(String internalName) {
     return _styleElementList.firstWhereOrNull(
        (el) => _getStyleName(el) == internalName);
  }


  /// Safely extracts the style name from a <w:style> element's <w:name> child.
  String? _getStyleName(XmlElement styleEl) {
      return styleEl
          .findElements('name', namespace: ns['w'])
          .firstOrNull
          ?.getAttribute('val', namespace: ns['w']);
  }

   /// Safely extracts the WD_STYLE_TYPE from a <w:style> element's w:type attribute.
   /// Defaults to PARAGRAPH if attribute is missing.
   WD_STYLE_TYPE _getStyleType(XmlElement styleEl) { // Return non-nullable, assume default
       final typeStr = styleEl.getAttribute('type', namespace: ns['w']);
       // Default to PARAGRAPH if type attribute is missing, as per OOXML spec
       return typeStr == null ? WD_STYLE_TYPE.PARAGRAPH
                              : WD_STYLE_TYPE.fromXml(typeStr); // Use enum's fromXml
   }

    /// Safely checks if a <w:style> element has w:default="1".
   bool _isDefaultStyle(XmlElement styleEl) {
       final defaultStr = styleEl.getAttribute('default', namespace: ns['w']);
       // Check against '1' or 'true' (case-insensitive)
       return defaultStr == '1' || defaultStr?.toLowerCase() == 'true';
   }
}
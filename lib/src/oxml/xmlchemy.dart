/// Path: lib/src/oxml/xmlchemy.dart
/// Based on python-docx: docx/oxml/xmlchemy.py
///
/// Base class and helpers for mapping Dart classes to Open XML elements.

// ignore_for_file: cascade_invocations

import 'package:collection/collection.dart';
import 'package:xml/xml.dart';
import 'exceptions.dart';
import 'ns.dart';
import 'simpletypes.dart';

// Note: The Python version uses metaclasses and descriptors extensively to
// dynamically add properties and methods based on declarations like
// `prop = ZeroOrOne(...)` or `attr = OptionalAttribute(...)`.
// Dart does not have direct equivalents for these mechanisms.
//
// Therefore, this translation provides a BaseOxmlElement class with helper
// methods. Subclasses inheriting from BaseOxmlElement in Dart will need to
// explicitly define their properties (getters/setters) and methods (like
// getOrAddChild, addChild, removeChild) by calling these helper methods.
// The declarative Python definitions are replaced by explicit implementation
// in Dart subclasses.

/// Represents a qualified XML name (Namespace URI + Local Name).
/// Used internally for clarity when working with package:xml.
class _QName {
  final String? namespaceUri;
  final String localName;

  _QName(this.namespaceUri, this.localName);

  /// Creates a _QName from a prefixed string like "w:p".
  factory _QName.fromQualifiedName(String qualifiedName) {
    if (qualifiedName.startsWith('{') && qualifiedName.contains('}')) {
      final closeIndex = qualifiedName.indexOf('}');
      if (closeIndex <= 1 || closeIndex >= qualifiedName.length - 1) {
        throw ArgumentError(
          "Qualified name '$qualifiedName' must include a non-empty namespace URI and local name",
        );
      }
      final uri = qualifiedName.substring(1, closeIndex);
      final local = qualifiedName.substring(closeIndex + 1);
      return _QName(uri.isEmpty ? null : uri, local);
    }
    final parts = qualifiedName.split(':');
    if (parts.length == 1) {
      // No prefix, assume no namespace (should be rare in OpenXML)
      return _QName(null, parts[0]);
    }
    if (parts.length != 2 || parts[0].isEmpty || parts[1].isEmpty) {
      throw ArgumentError(
        "Qualified name '$qualifiedName' must be in the form 'prefix:localName'",
      );
    }
    final prefix = parts[0];
    final localName = parts[1];
    final uri = nsmap[prefix];
    if (uri == null) {
      // Allow 'xml:' prefix even if not explicitly in nsmap
      if (prefix == 'xml') {
        return _QName('http://www.w3.org/XML/1998/namespace', localName);
      }
      throw ArgumentError("Namespace prefix '$prefix' not found in nsmap");
    }
    return _QName(uri, localName);
  }

  /// Returns the fully qualified name in Clark notation "{uri}localName".
  String get qName =>
      namespaceUri == null ? localName : '{$namespaceUri}$localName';
}

/// Base class for all custom OXML element classes. Wraps an [XmlElement].
///
/// Provides common OxmL helper methods. Subclasses should explicitly define
/// properties and methods for accessing attributes and child elements,
/// using the helper methods provided here.
class BaseOxmlElement {
  final XmlElement element;

  /// Creates a wrapper around an existing [XmlElement].
  BaseOxmlElement(this.element);

  /// Returns the local part of a qualified tag name string.
  /// Example: qnName("{http://...}p") => "p"
  ///          qnName("w:p")         => "p"
  /// Needed by descriptors when throwing errors with user-friendly names.
  String qnName(String qualifiedTagName) {
    if (qualifiedTagName.contains(':')) {
      return qualifiedTagName.split(':')[1];
    } else if (qualifiedTagName.startsWith('{') &&
        qualifiedTagName.contains('}')) {
      return qualifiedTagName.substring(qualifiedTagName.indexOf('}') + 1);
    }
    return qualifiedTagName; // Assume it's already local if no prefix/URI
  }

  T? getParentAs<T extends BaseOxmlElement>(
      T Function(XmlElement) constructor) {
    final parentNode = element.parent;
    if (parentNode is XmlElement) {
      return constructor(parentNode);
    }
    return null;
  }

  /// Finds all direct child elements matching [targetQnTagName] and constructs
  /// instances of type [T] using the provided [constructor].
  ///
  /// [targetQnTagName] should be the prefixed tag name like "w:r".
  /// [constructor] should be a function that takes an [XmlElement] and returns
  /// an instance of [T], e.g., `(el) => CT_R(el)`.
  List<T> childrenWhereType<T extends BaseOxmlElement>(
    String targetQnTagName,
    T Function(XmlElement) constructor,
  ) {
    final qName = _QName.fromQualifiedName(targetQnTagName);
    return element.children
        .whereType<XmlElement>() // Filtra apenas elementos XML
        .where((el) => // Filtra pelo nome local e namespace URI corretos
            el.name.local == qName.localName &&
            el.name.namespaceUri == qName.namespaceUri)
        .map((el) => constructor(el)) // Constrói o objeto Dart wrapper
        .toList();
  }

  /// Returns the first child element matching one of the prefixed tag names
  /// in [tagnames] (e.g., ["w:p", "w:tbl"]).
  /// Returns `null` if no matching child is found.
  XmlElement? firstChildFoundIn(List<String> tagnames) {
    for (final tagname in tagnames) {
      final qName = _QName.fromQualifiedName(tagname);
      final child = element
          .findElements(qName.localName, namespace: qName.namespaceUri)
          .firstOrNull;
      if (child != null) {
        return child;
      }
    }
    return null;
  }

  /// Inserts [child] before the first child element matching one of the
  /// prefixed tag names in [successorTagnames].
  /// If no successor is found, [child] is appended.
  /// Returns the inserted [child] element.
  XmlElement insertElementBefore(
      XmlElement child, List<String> successorTagnames) {
    XmlElement? successor;
    for (final tagname in successorTagnames) {
      final qName = _QName.fromQualifiedName(tagname);
      // Find among direct children only
      successor = element.children
          .whereType<XmlElement>() // Consider only elements
          .firstWhereOrNull(
            (el) =>
                el.name.local == qName.localName &&
                el.name.namespaceUri == qName.namespaceUri,
          );
      if (successor != null) break;
    }

    if (successor != null) {
      final index = element.children.indexOf(successor);
      if (index != -1) {
        // Insert the new child at that index
        // package:xml handles setting the parent automatically here
        element.children.insert(index, child);
      } else {
        // Should not happen if found above, but fallback to append
        element.children.add(child); // parent is set automatically
        print(
          'Warn: Successor element ${successor.name.qualified} found but not in direct children list? Appending ${child.name.qualified} instead.',
        );
      }
    } else {
      // No successor found, append to the end
      element.children.add(child); // parent is set automatically
    }
    // DO NOT manually set child.parent = element; it's handled by the list operations.
    return child;
  }

  /// Removes all child elements matching any of the prefixed tag names
  /// in [tagnames] (e.g., ["w:p", "w:rPr"]).
  void removeAll(List<String> tagnames) {
    final childrenToRemove = <XmlNode>[];
    for (final tagname in tagnames) {
      final qName = _QName.fromQualifiedName(tagname);
      childrenToRemove.addAll(
          element.findElements(qName.localName, namespace: qName.namespaceUri));
    }
    // Remove in reverse order or use a copy to avoid concurrent modification issues
    for (final child in childrenToRemove.reversed) {
      element.children.remove(child);
    }
  }

  /// Returns the XML string for this element, pretty-printed. No XML declaration.
  /// Suitable for testing or debugging.
  String get xml {
    return element.toXmlString(pretty: true, indent: '  ');
  }

  /// The namespace-prefixed tag name of this element, e.g., "w:p".
  /// Returns local name if the namespace is not found in `pfxmap`.
  String get nsptag {
    final uri = element.name.namespaceUri;
    final prefix = uri != null ? pfxmap[uri] : null;
    if (prefix != null) {
      return '$prefix:${element.name.local}';
    }
    return element.name.local; // Fallback if namespace not found
  }

  @override
  String toString() {
    return "<${runtimeType} '<$nsptag>' at ${identityHashCode(this)}>";
  }

  // --- Helper methods for subclasses to implement properties ---
  // These replace the Python descriptor functionality.

  /// Helper to get an optional attribute value, converting it using [st].
  /// Returns [defaultValue] if the attribute is not present or conversion fails.
  /// [attrName] should be like "w:val" or just "val".
  T? getAttrVal<T>(String attrName, BaseSimpleType<T> st, {T? defaultValue}) {
    final qName = _QName.fromQualifiedName(attrName);
    final attrValue =
        element.getAttribute(qName.localName, namespace: qName.namespaceUri);
    if (attrValue == null) {
      return defaultValue;
    }
    try {
      return st.fromXml(attrValue);
    } catch (e) {
      // Log or handle error appropriately
      print(
        'Warn: Error converting attribute $attrName value "$attrValue": $e. Returning default.',
      );
      return defaultValue;
    }
  }

  /// Helper to set an optional attribute value, converting it using [st].
  /// Removes the attribute if [value] is null or equals [defaultValue].
  /// [attrName] should be like "w:val" or just "val".
  void setAttrVal<T>(String attrName, T? value, BaseSimpleType<T> st,
      {T? defaultValue}) {
    final qName = _QName.fromQualifiedName(attrName);
    final nsUri = qName.namespaceUri;
    final localName = qName.localName;

    // Check against null first, then default value if provided
    bool shouldRemove = (value == null);
    if (!shouldRemove && defaultValue != null && value == defaultValue) {
      shouldRemove = true;
    }

    if (shouldRemove) {
      element.removeAttribute(localName, namespace: nsUri);
      return;
    }

    // If value is not null and not the default, proceed with conversion and setting
    if (value != null) {
      try {
        final xmlValue = st.toXml(value);
        // toXml might return null for specific enum cases (like INHERITED)
        if (xmlValue == null) {
          element.removeAttribute(localName, namespace: nsUri);
        } else {
          element.setAttribute(localName, xmlValue, namespace: nsUri);
        }
      } catch (e) {
        // Log or handle error appropriately
        print(
          'Warn: Error converting attribute $attrName value "$value" to XML: $e. Attribute not set.',
        );
      }
    }
  }

  /// Helper to get a required attribute value, converting it using [st].
  /// Throws [InvalidXmlError] if the attribute is not present.
  /// Throws conversion errors if conversion fails.
  /// [attrName] should be like "w:val" or just "val".
  T getReqAttrVal<T>(String attrName, BaseSimpleType<T> st) {
    final qName = _QName.fromQualifiedName(attrName);
    final attrValue =
        element.getAttribute(qName.localName, namespace: qName.namespaceUri);
    if (attrValue == null) {
      throw InvalidXmlError(
        "required '$attrName' attribute not present on element ${element.name.qualified}",
      );
    }
    try {
      return st.fromXml(attrValue);
    } catch (e) {
      print(
        'Error converting required attribute $attrName value "$attrValue": $e',
      );
      rethrow; // Re-throw as it's required
    }
  }

  /// Helper to set a required attribute value, converting it using [st].
  /// Throws [ArgumentError] if [value] is null or conversion results in null.
  /// Throws conversion errors if conversion fails.
  /// [attrName] should be like "w:val" or just "val".
  void setReqAttrVal<T>(String attrName, T value, BaseSimpleType<T> st) {
    String? xmlValue;
    try {
      xmlValue = st.toXml(value);
    } catch (e) {
      print(
          'Error converting required attribute $attrName value "$value" to XML: $e');
      rethrow;
    }

    if (xmlValue == null) {
      // Should not happen for required attributes unless toXml logic is flawed
      throw ArgumentError(
          "value '$value' converted to null for required attribute '$attrName'");
    }
    final qName = _QName.fromQualifiedName(attrName);
    element.setAttribute(qName.localName, xmlValue,
        namespace: qName.namespaceUri);
  }

  /// Finds the first child element matching the prefixed tag name [tagName].
  /// Returns null if not found.
  XmlElement? childOrNull(String tagName) {
    final qName = _QName.fromQualifiedName(tagName);
    return element
        .findElements(qName.localName, namespace: qName.namespaceUri)
        .firstOrNull;
  }

  /// Finds all child elements matching the prefixed tag name [tagName].
  List<XmlElement> children(String tagName) {
    final qName = _QName.fromQualifiedName(tagName);
    return element
        .findElements(qName.localName, namespace: qName.namespaceUri)
        .toList();
  }

  /// Gets the child element matching [tagName] or creates and adds it if not present.
  /// [newChildFactory] should be a function that returns a new instance of the
  /// desired child element (e.g., `() => XmlElement(XmlName('p', nsmap['w']))`).
  /// [successors] is a list of prefixed tag names that the new element should
  /// be inserted before.
  XmlElement getOrAddChild(String tagName, List<String> successors,
      XmlElement Function() newChildFactory) {
    var child = childOrNull(tagName);
    if (child == null) {
      child = newChildFactory();
      insertElementBefore(child, successors);
    }
    return child;
  }

  /// Removes the child element matching the prefixed tag name [tagName] if it exists.
  void removeChild(String tagName) {
    removeAll([tagName]);
  }

  /// Creates a new 'loose' element using the provided factory function.
  /// The factory should return the desired new element.
  XmlElement newElement(XmlElement Function() factory) {
    return factory();
  }

  /// Inserts a child element in the correct sequence according to successors.
  /// Returns the inserted child.
  XmlElement insertChild(XmlElement child, List<String> successors) {
    return insertElementBefore(child, successors);
  }

  /// Creates a new element using the factory and inserts it according to successors.
  /// Returns the newly added child.
  XmlElement addChildElement(
      List<String> successors, XmlElement Function() factory) {
    final child = factory();
    return insertElementBefore(child, successors);
  }

  /// Finds the first required child element matching the prefixed tag name [tagName].
  /// Throws [InvalidXmlError] if not found.
  XmlElement oneAndOnlyOneChild(String tagName) {
    final child = childOrNull(tagName);
    if (child == null) {
      throw InvalidXmlError(
        "required '$tagName' child element not present in ${element.name.qualified}",
      );
    }
    return child;
  }

  /// Finds the required list of child elements matching the prefixed tag name [tagName].
  /// Throws [InvalidXmlError] if the list is empty.
  List<XmlElement> oneOrMoreChildren(String tagName) {
    final childrenV = children(tagName);
    if (childrenV.isEmpty) {
      throw InvalidXmlError(
        "required '$tagName' child element not present (minimum one required) in ${element.name.qualified}",
      );
    }
    return childrenV;
  }

  // --- Choice Group Helpers ---

  /// Gets the currently active element from a choice group specified by
  /// the list of prefixed tag names [choiceTagnames]. Returns null if none are present.
  XmlElement? getChoiceElement(List<String> choiceTagnames) {
    return firstChildFoundIn(choiceTagnames);
  }

  /// Removes any existing element belonging to the choice group specified by
  /// the list of prefixed tag names [choiceTagnames].
  void removeChoiceGroup(List<String> choiceTagnames) {
    removeAll(choiceTagnames);
  }

  /// Gets the specific choice element [choiceTagName], replacing any other
  /// element from the group [allChoiceTagnames] if necessary.
  /// [newChoiceFactory] creates the new element if needed.
  /// [successors] defines the insertion order.
  XmlElement getOrChangeToChoice(
    String choiceTagName,
    List<String> allChoiceTagnames,
    List<String> successors,
    XmlElement Function() newChoiceFactory,
  ) {
    final qNameChoice = _QName.fromQualifiedName(choiceTagName);
    XmlElement? currentChoice = getChoiceElement(allChoiceTagnames);

    if (currentChoice != null &&
        currentChoice.name.local == qNameChoice.localName &&
        currentChoice.name.namespaceUri == qNameChoice.namespaceUri) {
      return currentChoice; // Already the correct choice
    }

    removeChoiceGroup(
        allChoiceTagnames); // Remove any existing different choice
    return addChildElement(successors, newChoiceFactory); // Add the new choice
  }
}

/// Serializes an [XmlElement] to a human-readable XML string suitable for tests.
/// No XML declaration.
String serializeForReading(XmlElement element) {
  return element.toXmlString(pretty: true, indent: '  ');
}

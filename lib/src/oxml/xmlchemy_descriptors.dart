/// File: xmlchemy_descriptors.dart
/// Defines descriptor classes for accessing child elements of BaseOxmlElement subclasses.

import 'package:xml/xml.dart';
//import 'ns.dart' show qn;
import 'xmlchemy.dart' show BaseOxmlElement;

/// Descriptor for ZERO or ONE child element of type [T] with qualified tag name [qnTagName].
class ZeroOrOne<T extends BaseOxmlElement> {
  /// The qualified tag name (e.g., "{namespace}tag").
  final String qnTagName;

  /// Optional sequence of qualified tag names of succeeding elements, used for insertion order.
  final List<String> successors;

  /// Creates a descriptor for an optional single child element.
  /// [qnTagName] is the qualified name like "{nsURI}localName".
  /// [successors] lists qualified names of elements that should come after this one.
  const ZeroOrOne(this.qnTagName, {this.successors = const []});

  /// Returns the wrapped child element [T] if found, otherwise `null`.
  /// [parent] is the instance of the owning BaseOxmlElement subclass.
  /// [factoryFn] is a function that takes an XmlElement and returns an instance of T.
  T? getElement(BaseOxmlElement parent, T Function(XmlElement) factoryFn) {
    final element = parent.childOrNull(qnTagName);
    return element != null ? factoryFn(element) : null;
  }

  /// Returns the wrapped child element [T], creating and inserting it if not found.
  /// [parent] is the instance of the owning BaseOxmlElement subclass.
  /// [factory] is a zero-argument function that returns a new, empty [XmlElement] for T.
  T getOrAdd(BaseOxmlElement parent, XmlElement Function() factory,
      T Function(XmlElement) factoryFn) {
    final element = parent.getOrAddChild(qnTagName, successors, factory);
    return factoryFn(element);
  }

  /// Removes all child elements matching [qnTagName] from the [parent].
  void remove(BaseOxmlElement parent) {
    parent.removeAll([qnTagName]);
  }
}

/// Descriptor for ZERO or MORE child elements of type [T] with qualified tag name [qnTagName].
class ZeroOrMore<T extends BaseOxmlElement> {
  /// The qualified tag name (e.g., "{namespace}tag").
  final String qnTagName;

  /// Function to create a wrapper instance [T] from an [XmlElement].
  final T Function(XmlElement) factoryFn;

  /// Creates a descriptor for an optional, repeatable child element.
  const ZeroOrMore(this.qnTagName, this.factoryFn);

  /// Returns a list of all wrapped child elements [T] found in [parent].
  List<T> getElements(BaseOxmlElement parent) {
    return parent.childrenWhereType<T>(qnTagName, factoryFn);
  }

  /// Removes all child elements matching [qnTagName] from the [parent].
  void removeAll(BaseOxmlElement parent) {
    parent.removeAll([qnTagName]);
  }

  // Note: Adding new elements (insertNew/addNew) is typically handled by
  // methods on the parent class (e.g., CT_Row.addTc()) rather than the descriptor.
  // The descriptor's main job is retrieving existing elements.
}

/// Descriptor for ONE and ONLY ONE child element of type [T] with qualified tag name [qnTagName].
class OneAndOnlyOne<T extends BaseOxmlElement> {
  /// The qualified tag name (e.g., "{namespace}tag").
  final String qnTagName;

  /// Function to create a wrapper instance [T] from an [XmlElement].
  final T Function(XmlElement) factoryFn;

  /// Creates a descriptor for a required single child element.
  const OneAndOnlyOne(this.qnTagName, this.factoryFn);

  /// Returns the wrapped child element [T].
  /// Throws [StateError] if the element is not found or if multiple are found
  /// (though `childOrNull` used in `BaseOxmlElement` typically returns the first).
  T getElement(BaseOxmlElement parent) {
    final element = parent.childOrNull(qnTagName);
    if (element == null) {
      throw StateError(
          'Required child element <${parent.qnName(qnTagName)}> not found');
    }
    // Consider adding check for multiple elements if strictness is required
    // final elements = parent.element.findElements(parent.qnName(qnTagName));
    // if (elements.length > 1) {
    //   print('Warning: Multiple required elements <${parent.qnName(qnTagName)}> found, using first.');
    // }
    return factoryFn(element);
  }

  // Removing a required element usually doesn't make sense, but could be added.
  // void remove(BaseOxmlElement parent) { ... }
}

/// Descriptor for ONE or MORE child elements of type [T] with qualified tag name [qnTagName].
class OneOrMore<T extends BaseOxmlElement> {
  /// The qualified tag name (e.g., "{namespace}tag").
  final String qnTagName;

  /// Function to create a wrapper instance [T] from an [XmlElement].
  final T Function(XmlElement) factoryFn;

  /// Creates a descriptor for a required, repeatable child element.
  const OneOrMore(this.qnTagName, this.factoryFn);

  /// Returns a list of all wrapped child elements [T] found in [parent].
  /// Throws [StateError] if no elements are found.
  List<T> getElements(BaseOxmlElement parent) {
    final elements = parent.childrenWhereType<T>(qnTagName, factoryFn);
    if (elements.isEmpty) {
      throw StateError(
          'Required child element(s) <${parent.qnName(qnTagName)}> not found');
    }
    return elements;
  }

  /// Removes all child elements matching [qnTagName] from the [parent].
  /// Note: This might leave the parent in an invalid state according to the schema.
  void removeAll(BaseOxmlElement parent) {
    parent.removeAll([qnTagName]);
  }
}

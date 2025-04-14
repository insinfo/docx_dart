/// Path: lib/src/shared.dart
/// Based on python-docx: docx/shared.py
///
/// Shared objects and base types used across docx modules.

import 'package:meta/meta.dart'; // For @internal

// Import necessary types. Use placeholder types if the actual classes aren't defined yet.
// It's better to import the actual types when available.
import 'opc/part.dart' show XmlPart;
import 'oxml/xmlchemy.dart' show BaseOxmlElement;
import 'parts/story.dart' show StoryPart;
import 'types.dart'; // Assuming ProvidesXmlPart, ProvidesStoryPart are here

/// Base class for length units like Inches, Cm, Mm, Pt, and Emu.
///
/// Behaves as an immutable integer count of English Metric Units (EMU),
/// 914,400 per inch. Provides convenience unit conversion methods via getters.
@immutable
class Length implements Comparable<Length> {
  static const int _emusPerInch = 914400;
  static const int _emusPerCm = 360000;
  static const int _emusPerMm = 36000;
  static const int _emusPerPt = 12700;
  static const int _emusPerTwip = 635;

  /// The length value in English Metric Units (EMU).
  final int emu;

  /// Creates a Length object with the specified EMU value.
  const Length(this.emu);

  /// The equivalent length expressed in centimeters (double).
  double get cm => emu / _emusPerCm;

  /// The equivalent length expressed in inches (double).
  double get inches => emu / _emusPerInch;

  /// The equivalent length expressed in millimeters (double).
  double get mm => emu / _emusPerMm;

  /// The equivalent length expressed in points (double).
  double get pt => emu / _emusPerPt;

  /// The equivalent length expressed in twips (integer).
  /// A twip is 1/20th of a point.
  int get twips => (emu / _emusPerTwip).round();

  /// Returns the raw EMU value. Useful for direct comparison or use as int.
  int call() => emu;

  // --- Comparison and Operators ---

  @override
  int compareTo(Length other) => emu.compareTo(other.emu);

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Length && runtimeType == other.runtimeType && emu == other.emu;

  bool operator <(Length other) => emu < other.emu;
  bool operator <=(Length other) => emu <= other.emu;
  bool operator >(Length other) => emu > other.emu;
  bool operator >=(Length other) => emu >= other.emu;

  /// Adds two Length objects.
  Length operator +(Length other) => Length(emu + other.emu);

  /// Subtracts another Length object.
  Length operator -(Length other) => Length(emu - other.emu);

  /// Negates the Length value (useful for inverse indents etc.).
  Length operator -() => Length(-emu);

  /// Multiplies Length by an integer or double scalar.
  Length operator *(num factor) => Length((emu * factor).round());

  /// Divides Length by an integer or double scalar.
  Length operator /(num divisor) {
    if (divisor == 0) throw ArgumentError('Division by zero');
    return Length((emu / divisor).round());
  }

  /// Integer division.
  Length operator ~/(num divisor) {
    if (divisor == 0) throw ArgumentError('Division by zero');
    return Length(emu ~/ divisor);
  }

  @override
  int get hashCode => emu.hashCode;

  @override
  String toString() => 'Length($emu EMU)'; // Or provide a default unit?
}

/// Convenience class for creating a [Length] in inches.
class Inches extends Length {
  /// Creates a Length object representing the given value in inches.
  Inches(double inches) : super((inches * Length._emusPerInch).round());
}

/// Convenience class for creating a [Length] in centimeters.
class Cm extends Length {
  /// Creates a Length object representing the given value in centimeters.
  Cm(double cm) : super((cm * Length._emusPerCm).round());
}

/// Convenience class for creating a [Length] directly from EMU.
class Emu extends Length {
  /// Creates a Length object representing the given value in EMU.
  Emu(int emu) : super(emu);
}

/// Convenience class for creating a [Length] in millimeters.
class Mm extends Length {
  /// Creates a Length object representing the given value in millimeters.
  Mm(double mm) : super((mm * Length._emusPerMm).round());
}

/// Convenience class for creating a [Length] in points (1/72 inch).
class Pt extends Length {
  /// Creates a Length object representing the given value in points.
  Pt(double points) : super((points * Length._emusPerPt).round());
}

/// Convenience class for creating a [Length] in twips (1/20 point).
class Twips extends Length {
  /// Creates a Length object representing the given value in twips.
  Twips(num twips) : super((twips * Length._emusPerTwip).round());
}

/// Immutable value object defining a particular RGB color.
@immutable
class RGBColor {
  final int r;
  final int g;
  final int b;

  /// Creates an RGBColor instance. Values must be between 0 and 255.
  const RGBColor(this.r, this.g, this.b)
      : assert(r >= 0 && r <= 255, 'r must be between 0 and 255'),
        assert(g >= 0 && g <= 255, 'g must be between 0 and 255'),
        assert(b >= 0 && b <= 255, 'b must be between 0 and 255');

  /// Creates an RGBColor instance from a hex string like "3C2F80" or "3c2f80".
  /// Throws [FormatException] if the string is invalid.
  factory RGBColor.fromString(String rgbHexString) {
    if (rgbHexString.length != 6 ||
        !RegExp(r'^[0-9a-fA-F]+$').hasMatch(rgbHexString)) {
      throw FormatException(
          "Invalid RGB hex string format: '$rgbHexString'. Expected 6 hex digits.");
    }
    try {
      final r = int.parse(rgbHexString.substring(0, 2), radix: 16);
      final g = int.parse(rgbHexString.substring(2, 4), radix: 16);
      final b = int.parse(rgbHexString.substring(4, 6), radix: 16);
      return RGBColor(r, g, b);
    } catch (e) {
      // Catch potential parsing errors even with regex check
      throw FormatException("Error parsing RGB hex string '$rgbHexString': $e");
    }
  }

  /// Returns the hex string representation (e.g., "3C2F80"). Uppercase.
  @override
  String toString() => '${r.toRadixString(16).toUpperCase().padLeft(2, '0')}'
      '${g.toRadixString(16).toUpperCase().padLeft(2, '0')}'
      '${b.toRadixString(16).toUpperCase().padLeft(2, '0')}';

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is RGBColor &&
          runtimeType == other.runtimeType &&
          r == other.r &&
          g == other.g &&
          b == other.b;

  @override
  int get hashCode => r.hashCode ^ g.hashCode ^ b.hashCode;
}

// lazyproperty: Dart doesn't have a direct decorator equivalent.
// The common pattern is to use a private nullable field and a public getter:
//
// class MyClass {
//   MyExpensiveObject? _lazyObject;
//
//   MyExpensiveObject get lazyObject {
//     return _lazyObject ??= _computeExpensiveObject();
//   }
//
//   MyExpensiveObject _computeExpensiveObject() {
//     // ... computation ...
//     return result;
//   }
// }
// Since it's a pattern rather than a class, it's not included here directly.

// write_only_property: Not a standard Dart concept. Use a setter without a getter.
// void set writeOnlyValue(String value) { /* ... */ }

/// Base class for proxy objects wrapping an underlying Open XML element.
/// Provides access to the element and its containing part.
class ElementProxy {
  @internal
  final BaseOxmlElement element; // Renamed _element to element for clarity
  final ProvidesXmlPart? _parent; // Can be null if proxy wraps root/unparented

  /// Wraps [element]. [parent] provides access to the part.
  ElementProxy(this.element, [this._parent]);

  /// The package part containing this object's element.
  /// Throws [StateError] if the parent context is not available.
  XmlPart get part {
    final parent = _parent; // Use local variable for null safety check
    if (parent == null) {
      throw StateError(
          "part is not accessible from this element proxy (no parent context)");
    }
    return parent.part;
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is ElementProxy &&
          runtimeType == other.runtimeType &&
          element == other.element; // Compare based on wrapped element

  @override
  int get hashCode => element.hashCode;
}

/// Base class for objects that have a parent providing access to the part.
/// Used for elements below the part level that might need part services.
class Parented {
  @internal
  final ProvidesXmlPart parent; // Renamed _parent

  Parented(this.parent);

  /// The package part containing this object.
  XmlPart get part => parent.part;
}

/// Base class for document elements residing within a story part (like DocumentPart, HeaderPart).
/// Provides access to the containing [StoryPart].
class StoryChild {
  @internal
  final ProvidesStoryPart parent; // Renamed _parent

  StoryChild(this.parent);

  /// The story part (e.g., DocumentPart, HeaderPart) containing this object.
  StoryPart get part => parent.part;
}

/// Accumulates string fragments and joins them with a separator on demand.
class TextAccumulator {
  final String _separator;
  final List<String> _texts = [];

  /// Creates an accumulator joining fragments with [_separator] (defaults to empty string).
  TextAccumulator([this._separator = ""]);

  /// Adds a text fragment to the accumulator.
  void push(String text) {
    _texts.add(text);
  }

  /// Returns the accumulated text joined by the separator and clears the buffer.
  /// Returns `null` if no text has been accumulated.
  List<String> pop() { // Changed return type to List<String>
    if (_texts.isEmpty) {
      return []; // Return empty list instead of null
    }
    final text = _texts.join(_separator);
    _texts.clear();
    return [text]; // Return list containing the single string
  }

  /// Returns the accumulated text joined by the separator *without* clearing the buffer.
  /// Returns `null` if no text has been accumulated.
  String? peek() {
    if (_texts.isEmpty) {
      return null;
    }
    return _texts.join(_separator);
  }

  /// Checks if any text has been accumulated.
  bool get isEmpty => _texts.isEmpty;
  bool get isNotEmpty => _texts.isNotEmpty;
}

// Note: BlockItemElement is a Python TypeAlias referring to specific XML elements.
// In Dart, this might be represented by a common interface or base class if needed,
// or simply handled using `is` checks or pattern matching on the `XmlElement` type.
// A direct typedef isn't strictly necessary unless used extensively for type hinting.
// typedef BlockItemElement = XmlElement; // Or a more specific base class/interface

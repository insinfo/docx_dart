/// Path: lib/src/oxml/shared.dart
/// Based on python-docx: docx/oxml/shared.py
///
/// Objects shared by modules in the docx.oxml subpackage, typically simple
/// elements primarily containing a 'w:val' attribute.

import 'package:xml/xml.dart';

import 'ns.dart';
import 'parser.dart'; // Assuming OxmlElement factory is defined here
import 'simpletypes.dart';
import 'xmlchemy.dart';

/// Wrapper for elements like `<w:numId>`, `<w:ilvl>`, `<w:abstractNumId>`, etc.
/// Contains a text representation of a decimal number (e.g., 42) in its `w:val` attribute.
class CT_DecimalNumber extends BaseOxmlElement {
  /// Wraps the given [element].
  CT_DecimalNumber(super.element);

  /// The integer value of the `w:val` attribute.
  int get val => getReqAttrVal('w:val', stDecimalNumberConverter);
  set val(int value) => setReqAttrVal('w:val', value, stDecimalNumberConverter);

  /// newElement Returns a new `CT_DecimalNumber`-equivalent element having the specified
  /// tag name [nsptagname] and `w:val` attribute set to [val].
  static XmlElement create(String nsptagname, int val) {
    // Assume OxmlElement creates a basic XmlElement with the given tag and attributes
    return OxmlElement(nsptagname, attrs: {qn("w:val"): val.toString()});
  }
}

/// Wrapper for elements like `<w:b>`, `<w:i>`, `<w:caps>`, etc.
/// Contains a boolean-ish string in its `w:val` attribute (e.g., "1", "0", "true", "false", "on", "off").
/// If the `w:val` attribute is omitted, the effective value is `true`.
class CT_OnOff extends BaseOxmlElement {
  /// Wraps the given [element].
  CT_OnOff(super.element);

  static XmlElement create(String qnTagName) => OxmlElement(qnTagName);

  /// The boolean value of the `w:val` attribute. Defaults to `true` if the
  /// attribute is not present. Assigning `true` (the default) removes the
  /// attribute, assigning `false` sets it to "0".
  bool get val {
    // Use the helper, providing the default value.
    // The `!` asserts non-null because a default is always provided.
    return getAttrVal('w:val', stOnOffConverter, defaultValue: true)!;
  }

  set val(bool? value) {
    // If value is null or true (the default), remove the attribute.
    if (value == null || value == true) {
      setAttrVal('w:val', null, stOnOffConverter, defaultValue: true);
    } else {
      // Otherwise (value is false), set the attribute.
      setAttrVal('w:val', false, stOnOffConverter, defaultValue: true);
    }
  }
}

/// Wrapper for elements like `<w:pStyle>`, `<w:tblStyle>`, `<w:name>`, etc.
/// Contains a string value in its `w:val` attribute.
class CT_String extends BaseOxmlElement {
  /// Wraps the given [element].
  CT_String(super.element);

  /// The string value of the required `w:val` attribute.
  String get val => getReqAttrVal('w:val', stStringConverter);
  set val(String value) => setReqAttrVal('w:val', value, stStringConverter);

  /// newElement Returns a new `CT_String`-equivalent element with the specified tag name
  /// [nsptagname] and `w:val` attribute set to [val].
  static XmlElement create(String nsptagname, String val) {
    // Assume OxmlElement creates a basic XmlElement with the given tag and attributes
    return OxmlElement(nsptagname, attrs: {qn("w:val"): val});
  }
}

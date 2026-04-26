/// Path: lib/src/oxml/settings.dart
/// Based on python-docx: docx/oxml/settings.py
/// Custom element classes related to document settings (<w:settings>).
import 'package:docx_dart/src/oxml/shared.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:xml/xml.dart';
import 'ns.dart' show qn;
import 'parser.dart' show OxmlElement;
import 'shared.dart' show CT_OnOff;
import 'xmlchemy.dart' show BaseOxmlElement;
import 'xmlchemy_descriptors.dart';

// ignore_for_file: camel_case_types, unnecessary_this, unused_field

/// `<w:settings>` element, root element for the settings part.
class CT_Settings extends BaseOxmlElement {
  CT_Settings(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:settings');

  // --- Define sequence for insertion ---
  // Note: This sequence is very long. Only including relevant parts around
  // 'evenAndOddHeaders' for this example. Add others as needed.
  static final _tagSeq = [
    // ... many tags before ...
    qn("w:defaultTableStyle"),
    qn("w:updateFields"),
    qn("w:evenAndOddHeaders"), // Index 47 in the Python list (0-based)
    qn("w:bookFoldRevPrinting"),
    // ... many tags after ...
  ];

  // --- Descriptor for ZeroOrOne child ---
  static final _evenAndOddHeaders =
      ZeroOrOne<CT_OnOff>(qn('w:evenAndOddHeaders'),
          // Successors start from the element *after* evenAndOddHeaders
          successors: _tagSeq.sublist(3));
  static final _updateFields =
      ZeroOrOne<CT_OnOff>(qn('w:updateFields'), successors: _tagSeq.sublist(2));

  // --- Property Accessors ---

  /// The `<w:evenAndOddHeaders>` child element, or `null` if not present.
  CT_OnOff? get evenAndOddHeaders =>
      _evenAndOddHeaders.getElement(this, (el) => CT_OnOff(el));
  CT_OnOff? get updateFields =>
      _updateFields.getElement(this, (el) => CT_OnOff(el));

  /// Returns the `<w:evenAndOddHeaders>` child, creating it if necessary.
  CT_OnOff getOrAddEvenAndOddHeaders() => _evenAndOddHeaders.getOrAdd(this,
      () => CT_OnOff.create(qn('w:evenAndOddHeaders')), (el) => CT_OnOff(el));
  CT_OnOff getOrAddUpdateFields() => _updateFields.getOrAdd(
      this, () => CT_OnOff.create(qn('w:updateFields')), (el) => CT_OnOff(el));

  /// Removes the `<w:evenAndOddHeaders>` child element if present.
  void removeEvenAndOddHeaders() => _evenAndOddHeaders.remove(this);
  void removeUpdateFields() => _updateFields.remove(this);

  /// Alias for removeEvenAndOddHeaders to match Python internal method name hint
  void remove_evenAndOddHeaders() => removeEvenAndOddHeaders();

  /// Value of `w:evenAndOddHeaders`. Defaults to `false` if the element is not present.
  /// The presence of the element `<w:evenAndOddHeaders/>` means `true` unless
  /// its `w:val` attribute is explicitly "false" or "0".
  bool get evenAndOddHeadersVal {
    final element = evenAndOddHeaders;
    // If element doesn't exist, the property is off (false)
    if (element == null) {
      return false;
    }
    // If element exists, its value determines the state (CT_OnOff handles val)
    return element.val; // Assumes CT_OnOff.val defaults to true if attr absent
  }

  /// Sets the `w:evenAndOddHeaders` state.
  /// Setting to `true` ensures the element exists (usually with no `w:val` or `w:val="true"`).
  /// Setting to `false` or `null` removes the element.
  set evenAndOddHeadersVal(bool? value) {
    if (value == null || !value) {
      removeEvenAndOddHeaders();
    } else {
      getOrAddEvenAndOddHeaders().val = true;
    }
  }

  /// Alias for evenAndOddHeadersVal setter
  set evenAndOddHeaders_val(bool? value) => evenAndOddHeadersVal = value;

  /// Alias for evenAndOddHeadersVal getter
  bool get evenAndOddHeaders_val => evenAndOddHeadersVal;

  /// True when Word should update document fields when the file is opened.
  bool get updateFieldsVal => updateFields?.val ?? false;

  set updateFieldsVal(bool? value) {
    if (value == null || !value) {
      removeUpdateFields();
    } else {
      getOrAddUpdateFields().val = true;
    }
  }

  set updateFields_val(bool? value) => updateFieldsVal = value;
  bool get updateFields_val => updateFieldsVal;
}

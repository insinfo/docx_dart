/// Path: lib/src/oxml/drawing.dart
/// Based on python-docx: docx/oxml/drawing.py
///
/// Custom element classes for DrawingML-related elements like `<w:drawing>`.

import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/parser.dart';
import 'package:xml/xml.dart';

import 'shape.dart'; // Might need CT_Inline, CT_Anchor later
import 'xmlchemy.dart'; // For BaseOxmlElement

/// `<w:drawing>` element, containing a DrawingML object like a picture or chart.
class CT_Drawing extends BaseOxmlElement {
  /// Wraps the [element] in a [CT_Drawing] instance.
  CT_Drawing(super.element);

  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:drawing');

  // --- Getters for potential DrawingML content ---
  // These would need to be implemented based on the actual expected children
  // like <wp:inline> or <wp:anchor>. The Python code doesn't show these details.

  /// Returns the first `<wp:inline>` child element, or null if not present.
  CT_Inline? get inline {
    final inlineEl = childOrNull('wp:inline');
    return inlineEl == null ? null : CT_Inline(inlineEl);
  }

  /// Returns the first `<wp:anchor>` child element, or null if not present.
  CT_Anchor? get anchor {
    final anchorEl = childOrNull('wp:anchor');
    return anchorEl == null ? null : CT_Anchor(anchorEl);
  }

  // --- Methods to add content (example) ---

  /// Appends a `<wp:inline>` or `<wp:anchor>` element as a child.
  /// This replaces any existing inline or anchor element.
  void setDrawingContent(XmlElement inlineOrAnchor) {
    // Remove existing inline/anchor first to ensure only one is present
    removeAll(['wp:inline', 'wp:anchor']);
    element.children.add(inlineOrAnchor);
    // Ensure parent is set correctly by package:xml
  }

  // Add other methods or properties specific to <w:drawing> as needed.
  // For example, accessing specific properties within the inline/anchor
  // could be proxied here, although it might be cleaner to access them
  // via the .inline or .anchor properties.
}

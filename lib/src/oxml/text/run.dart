/// Path: lib/src/oxml/text/run.dart
/// Based on python-docx: docx/oxml/text/run.py
/// Custom element classes related to text runs (CT_R).

import 'package:xml/xml.dart';

import '../drawing.dart';
import '../ns.dart'; // qn, nsmap
import '../parser.dart'; // OxmlElement
import '../simpletypes.dart';
import '../xmlchemy.dart'; // BaseOxmlElement etc.
import '../../shared.dart'; // TextAccumulator
import 'font.dart';
import 'pagebreak.dart';
import 'parfmt.dart' show CT_TabStop;

// ----------------------------------------------------------------------------
// Run-level elements
// ----------------------------------------------------------------------------

/// `<w:r>` element, containing the properties and text for a run.
class CT_R extends BaseOxmlElement {
  CT_R(super.element);

  /// Static factory method to create a new minimal <w:r> element.
  static XmlElement create() => OxmlElement(qnTagName);

  static final qnTagName = qn('w:r');

  // --- Getters for child elements ---

  /// The optional `<w:rPr>` child element specifying run formatting properties.
  CT_RPr? get rPr {
    // Ensure CT_RPr.qnTagName is defined in font.dart
    final child = childOrNull(CT_RPr.qnTagName);
    return child != null ? CT_RPr(child) : null;
  }

  /// Sequence of `<w:br>` child elements (breaks).
  List<CT_Br> get brElements =>
      childrenWhereType<CT_Br>(CT_Br.qnTagName, (el) => CT_Br(el));

  /// Sequence of `<w:cr>` child elements (carriage returns).
  List<CT_Cr> get crElements =>
      childrenWhereType<CT_Cr>(CT_Cr.qnTagName, (el) => CT_Cr(el));

  /// Sequence of `<w:drawing>` child elements.
  List<CT_Drawing> get drawingElements => childrenWhereType<CT_Drawing>(
      CT_Drawing.qnTagName, (el) => CT_Drawing(el));

  /// Sequence of `<w:t>` child elements (text).
  List<CT_Text> get tElements =>
      childrenWhereType<CT_Text>(CT_Text.qnTagName, (el) => CT_Text(el));

  /// Sequence of `<w:tab/>` child elements (tab characters).
  /// Note: Uses CT_TabStop wrapper for convenience, leveraging its toString().
  List<CT_TabStop> get tabElements => childrenWhereType<CT_TabStop>(
      CT_TabStop.qnTagName, (el) => CT_TabStop(el));

  // --- Get or Add methods ---

  /// Return the `<w:rPr>` child element, creating it if not present and
  /// inserting it as the first child.
  CT_RPr getOrAddRPr() {
    var rPrElement = childOrNull(CT_RPr.qnTagName);
    if (rPrElement == null) {
      // Assumes CT_RPr.create() exists in font.dart
      rPrElement = CT_RPr.create();
      _insertRPr(CT_RPr(rPrElement)); // Use helper to insert at beginning
    }
    return CT_RPr(rPrElement);
  }

  // --- Add methods ---

  /// Add a new `<w:br>` child element and return its wrapper `CT_Br`.
  CT_Br addBr({String? type, String? clear}) {
    final brElement = CT_Br.create(type: type, clear: clear);
    // Append the new element using a base class helper or directly
    element.children.add(brElement);
    return CT_Br(brElement);
  }

  /// Add a new `<w:tab/>` child element and return its wrapper `CT_TabStop`.
  CT_TabStop addTab() {
    final tabElement = CT_TabStop.createRunTab();
    element.children.add(tabElement);
    return CT_TabStop(tabElement);
  }

  /// Add a new `<w:drawing>` child element (internal use).
  CT_Drawing _addDrawing() {
    // Assumes CT_Drawing.create() exists in drawing.dart
    final drawingElement = CT_Drawing.create();
    element.children.add(drawingElement);
    return CT_Drawing(drawingElement);
  }

  /// Add a new `<w:t>` child element with [text] (internal use).
  CT_Text _addT(String text) {
    final tElement = CT_Text.create(text: text);
    element.children.add(tElement);
    return CT_Text(tElement);
  }

  /// Return a newly added `<w:t>` element containing [text].
  /// Sets `xml:space="preserve"` if [text] contains leading/trailing whitespace.
  CT_Text addT(String text) {
    final t = _addT(text);
    // CT_Text.create handles xml:space
    return t;
  }

  /// Return newly appended `CT_Drawing` (`w:drawing`) child element.
  /// The `w:drawing` element has [inlineOrAnchor] (e.g., CT_Inline) as its child.
  CT_Drawing addDrawing(BaseOxmlElement inlineOrAnchor) {
    final drawing = _addDrawing();
    drawing.element.children.add(inlineOrAnchor.element.copy());
    return drawing;
  }

  /// Remove all child elements except a `w:rPr` element if present.
  void clearContent() {
    final rPrElement = childOrNull(CT_RPr.qnTagName);
    element.children.clear();
    if (rPrElement != null) {
      element.children.add(rPrElement);
    }
  }

  /// Yields the inner content items (String for text, CT_Drawing,
  /// CT_LastRenderedPageBreak) in document order. Contiguous text runs are merged.
  List<dynamic /* String | CT_Drawing | CT_LastRenderedPageBreak */ >
      get innerContentItems {
    final accum = TextAccumulator();
    final items = <dynamic>[];

    final relevantTextTags = {
      CT_Br.qnTagName,
      CT_Cr.qnTagName,
      CT_NoBreakHyphen.qnTagName,
      CT_PTab.qnTagName,
      CT_Text.qnTagName,
      CT_TabStop.qnTagName, // For <w:tab/>
    };

    // Iterate safely over element.children which is non-nullable
    for (final child in element.children) {
      if (child is! XmlElement) continue; // Skip non-element nodes

      final qName = child.name.qualified;
      dynamic itemToAdd;
      String? textEquivalent;

      // Skip rPr
      if (qName == CT_RPr.qnTagName) continue;

      // Identify specific object types
      if (qName == CT_Drawing.qnTagName) {
        itemToAdd = CT_Drawing(child);
      } else if (qName == CT_LastRenderedPageBreak.qnTagName) {
        itemToAdd = CT_LastRenderedPageBreak(child);
      }
      // Get text equivalent for known text-contributing elements
      else if (relevantTextTags.contains(qName)) {
        // Use polymorphism via toString() defined in each CT_* class
        if (qName == CT_Br.qnTagName)
          textEquivalent = CT_Br(child).toString();
        else if (qName == CT_Cr.qnTagName)
          textEquivalent = CT_Cr(child).toString();
        else if (qName == CT_NoBreakHyphen.qnTagName)
          textEquivalent = CT_NoBreakHyphen(child).toString();
        else if (qName == CT_PTab.qnTagName)
          textEquivalent = CT_PTab(child).toString();
        else if (qName == CT_Text.qnTagName)
          textEquivalent = CT_Text(child).toString();
        else if (qName == CT_TabStop.qnTagName)
          textEquivalent = CT_TabStop(child).toString();
      }

      if (itemToAdd != null) {
        // Flush buffered text before adding a non-text item
        for (final text in accum.pop()) {
          items.add(text);
        }
        items.add(itemToAdd);
      } else if (textEquivalent != null) {
        accum.push(textEquivalent);
      }
      // Ignore unrecognized elements within the run
    }

    // Add any remaining text in the accumulator
    for (final text in accum.pop()) {
      items.add(text);
    }

    return items;
  }

  /// All `<w:lastRenderedPageBreak>` child elements of this run.
  List<CT_LastRenderedPageBreak> get lastRenderedPageBreaks =>
      childrenWhereType<CT_LastRenderedPageBreak>(
          CT_LastRenderedPageBreak.qnTagName,
          (el) => CT_LastRenderedPageBreak(el));

  /// The style ID string from the `w:val` attribute of the `./w:rPr/w:rStyle`
  /// grandchild. Returns `null` if no explicit style is applied.
  String? get style => rPr?.style;

  /// Set the character style of this run using its style ID [styleId].
  /// If [styleId] is null, removes the explicit character style.
  set style(String? styleId) {
    final rPr = getOrAddRPr();
    // Assumes CT_RPr has a 'style' setter that handles the <w:rStyle> child
    rPr.style = styleId;
  }

  /// Gets the concatenated text content of this run.
  /// Child elements like `<w:t>`, `<w:tab>`, `<w:br>`, `<w:cr>`, etc., are
  /// converted to their string representations (`text`, `\t`, `\n`).
  /// Elements like `<w:drawing>` are ignored.
  String get text {
    final buffer = StringBuffer();
    final wUri = nsmap['w']!; // Get WordprocessingML namespace URI once

    for (final child in element.children) {
      if (child is! XmlElement || child.name.namespaceUri != wUri) continue;

      final localName = child.name.local;

      // Use toString() polymorphism defined in each CT_* class
      if (localName == 't')
        buffer.write(CT_Text(child).toString());
      else if (localName == 'tab')
        buffer.write(CT_TabStop(child).toString());
      else if (localName == 'cr')
        buffer.write(CT_Cr(child).toString());
      else if (localName == 'br')
        buffer.write(CT_Br(child).toString());
      else if (localName == 'noBreakHyphen')
        buffer.write(CT_NoBreakHyphen(child).toString());
      else if (localName == 'ptab') buffer.write(CT_PTab(child).toString());
      // Other elements like <w:drawing>, <w:rPr> are ignored
    }
    return buffer.toString();
  }

  /// Replaces all content of this run (except `<w:rPr>`) with appropriate
  /// child elements corresponding to the characters in [text]. Handles tabs (`\t`)
  /// and line breaks (`\n`, `\r`) by creating `<w:tab/>` and `<w:br/>` elements.
  /// Preserves existing run formatting (`<w:rPr>`).
  set text(String text) {
    clearContent(); // Clears content but preserves rPr
    _RunContentAppender.appendToRunFromText(this, text);
  }

  /// Helper for inserting <w:rPr> at the beginning.
  CT_RPr _insertRPr(CT_RPr rPr) {
    // Ensure rPr element is valid before inserting
    if (rPr.element.parent != null) {
      rPr.element.parent!.children.remove(rPr.element); // Detach if necessary
    }
    element.children.insert(0, rPr.element);
    return rPr;
  }
}

// ----------------------------------------------------------------------------
// Run inner-content elements
// ----------------------------------------------------------------------------

/// `<w:br>` element, indicating a line, page, or column break in a run.
class CT_Br extends BaseOxmlElement {
  CT_Br(super.element);

  /// Break type like "page", "column", or "textWrapping" (default).
  String get type => // Changed to non-nullable, defaulting in getter
      getAttrVal('w:type', stBrTypeConverter, defaultValue: "textWrapping")!;
  set type(String? value) => // Allow null to set default
      setAttrVal('w:type', value, stBrTypeConverter,
          defaultValue: "textWrapping");

  /// Text wrapping break type like "left", "right", or "all". Optional.
  String? get clear => getAttrVal('w:clear', stBrClearConverter);
  set clear(String? value) => setAttrVal('w:clear', value, stBrClearConverter);

  /// Creates a new `XmlElement` suitable for a CT_Br.
  static XmlElement create({String? type, String? clear}) {
    final attrs = <String, String>{};
    // Only add attributes if they differ from the default or are explicitly set
    if (type != null && type != "textWrapping") {
      attrs[qn('w:type')] = stBrTypeConverter.toXml(type)!;
    }
    if (clear != null) {
      attrs[qn('w:clear')] = stBrClearConverter.toXml(clear)!;
    }
    return OxmlElement(qnTagName, attrs: attrs);
  }

  /// Text equivalent of this element. Actual value depends on break type.
  /// A line break ("textWrapping") or default is translated as "\n". Others are "".
  @override
  String toString() => (type == "textWrapping") ? "\n" : "";

  static final qnTagName = qn('w:br');
}

/// `<w:cr>` element, representing a carriage-return (0x0D) character.
/// NOTE: In XML, this is CT_Empty, but given a distinct class for behavior.
class CT_Cr extends BaseOxmlElement {
  CT_Cr(super.element);

  /// Creates a new `XmlElement` suitable for a CT_Cr.
  static XmlElement create() => OxmlElement(qnTagName);

  /// Text equivalent of this element, a single newline ("\n").
  @override
  String toString() => "\n";

  static final qnTagName = qn('w:cr');
}

/// `<w:noBreakHyphen>` element, a non-breaking hyphen character.
/// NOTE: In XML, this is CT_Empty, but given a distinct class for behavior.
class CT_NoBreakHyphen extends BaseOxmlElement {
  CT_NoBreakHyphen(super.element);

  /// Creates a new `XmlElement` suitable for a CT_NoBreakHyphen.
  static XmlElement create() => OxmlElement(qnTagName);

  /// Text equivalent of this element, a single dash character ("-").
  @override
  String toString() => "-";

  static final qnTagName = qn('w:noBreakHyphen');
}

/// `<w:ptab>` element, representing an absolute-position tab character.
/// NOTE: In XML schema, this element has specific attributes related to position.
/// Here, it's simplified like the Python version for text extraction.
class CT_PTab extends BaseOxmlElement {
  CT_PTab(super.element);

  /// Creates a new `XmlElement` suitable for a CT_PTab.
  static XmlElement create() => OxmlElement(qnTagName);

  /// Text equivalent of this element, a single tab ("\t") character.
  @override
  String toString() => "\t";

  static final qnTagName = qn('w:ptab');
}

// Note: CT_Tab (<w:tab/>) functionality for text runs is handled by CT_TabStop's toString()
// from the parfmt module, as it uses the same tag. Its qnTagName is defined there.

/// `<w:t>` element, containing a sequence of characters within a run.
class CT_Text extends BaseOxmlElement {
  CT_Text(super.element);

  /// Text content of the element. Returns empty string if no text node child exists.
  String get textValue => element.innerText; // innerText is usually sufficient

  /// Sets the text content of the element. Replaces any existing children.
  /// Manages `xml:space="preserve"` attribute.
  set textValue(String? value) {
    element.children.clear(); // Remove existing text nodes/children
    if (value != null && value.isNotEmpty) {
      element.children.add(XmlText(value));
      // Preserve whitespace if necessary
      if (value.trim().length < value.length) {
        element.setAttribute(qn('xml:space'), 'preserve');
      } else {
        element.removeAttribute(qn('xml:space')); // Remove if not needed
      }
    } else {
      element.removeAttribute(qn('xml:space')); // Remove if text is empty
    }
  }

  /// Creates a new `XmlElement` suitable for a CT_Text, containing the given text.
  /// Sets `xml:space="preserve"` if text has leading/trailing whitespace.
  static XmlElement create({String text = ''}) {
    final tElement = OxmlElement(qnTagName);
    if (text.isNotEmpty) {
      tElement.children.add(XmlText(text));
      if (text.trim().length < text.length) {
        tElement.setAttribute(qn('xml:space'), 'preserve');
      }
    }
    return tElement;
  }

  /// Text contained in this element, the empty string if it has no content.
  @override
  String toString() => textValue;

  static final qnTagName = qn('w:t');
}

// ----------------------------------------------------------------------------
// Utility
// ----------------------------------------------------------------------------

/// Translates a Dart string into run content elements appended in a `w:r` element.
/// Ported from Python's _RunContentAppender.
class _RunContentAppender {
  final CT_R _r;
  final StringBuffer _buffer = StringBuffer();

  _RunContentAppender(this._r);

  /// Append inner-content elements for [text] to the [r] element.
  static void appendToRunFromText(CT_R r, String text) {
    final appender = _RunContentAppender(r);
    appender.addText(text);
  }

  /// Append inner-content elements for [text] to the `w:r` element.
  void addText(String text) {
    for (int i = 0; i < text.length; i++) {
      this.addChar(text[i]);
    }
    this.flush(); // Flush any remaining buffer content at the end
  }

  /// Process next character of input. Appends special elements directly,
  /// buffers regular characters.
  void addChar(String char) {
    if (char == '\t') {
      this.flush(); // Write buffered text before adding tab
      this._r.addTab(); // Assumes CT_R.addTab() exists
    } else if (char == '\n' || char == '\r') {
      // Map both \n and \r to <w:br/> for line breaks within a run
      this.flush(); // Write buffered text before adding break
      this._r.addBr(); // Assumes CT_R.addBr() exists
    } else {
      // Buffer regular characters
      this._buffer.write(char);
    }
  }

  /// Write any pending text buffer to a new `<w:t>` element and clear buffer.
  void flush() {
    final text = this._buffer.toString();
    if (text.isNotEmpty) {
      this._r.addT(text); // CT_R.addT handles xml:space preservation
      this._buffer.clear();
    }
  }
}

// --- Placeholder Converters (Implementation assumed in simpletypes.dart) ---
class ST_BrTypeConverter implements BaseSimpleType<String> {
  const ST_BrTypeConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {} // Could add validation vs enum values
}

class ST_BrClearConverter implements BaseSimpleType<String> {
  const ST_BrClearConverter();
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;
  @override
  void validate(String value) {} // Could add validation vs enum values
}

// Instantiate placeholder converters
final stBrTypeConverter = const ST_BrTypeConverter();
final stBrClearConverter = const ST_BrClearConverter();

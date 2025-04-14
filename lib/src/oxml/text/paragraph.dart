/// Path: lib/src/oxml/text/paragraph.dart
/// Based on python-docx: docx/oxml/text/paragraph.py
/// Custom element classes related to paragraphs (CT_P).

import 'package:xml/xml.dart';

import '../../enum/text.dart' show WD_PARAGRAPH_ALIGNMENT;
import '../ns.dart'; // qn and namespaces
import '../parser.dart'; // OxmlElement
import '../section.dart'; // CT_SectPr
import '../xmlchemy.dart'; // BaseOxmlElement
import 'hyperlink.dart'; // CT_Hyperlink
import 'pagebreak.dart'; // CT_LastRenderedPageBreak
import 'parfmt.dart'; // CT_PPr
import 'run.dart'; // CT_R

class CT_P extends BaseOxmlElement {
  CT_P(super.element);

  // --- Child Element Access ---

  /// The `<w:pPr>` child element specifying paragraph formatting, or `null` if not present.
  CT_PPr? get pPr {
    final child = childOrNull(CT_PPr.qnTagName);
    return child != null ? CT_PPr(child) : null;
  }

  /// Sequence of `<w:hyperlink>` child elements.
  List<CT_Hyperlink> get hyperlinkElements => childrenWhereType<CT_Hyperlink>(
      CT_Hyperlink.qnTagName, (el) => CT_Hyperlink(el));

  /// Alias for `hyperlinkElements` to match Python's `hyperlink_lst`.
  List<CT_Hyperlink> get hyperlink_lst => hyperlinkElements;

  /// Sequence of `<w:r>` child elements.
  List<CT_R> get rElements =>
      childrenWhereType<CT_R>(CT_R.qnTagName, (el) => CT_R(el));

  /// Alias for `rElements` to match Python's `r_lst`.
  List<CT_R> get r_lst => rElements;

  // --- Get or Add Methods ---

  /// Return the `<w:pPr>` child element, creating it if not present and
  /// inserting it as the first child.
  CT_PPr getOrAddPPr() {
    var pPrElement = childOrNull(CT_PPr.qnTagName);
    if (pPrElement == null) {
      // Assumes CT_PPr.create() exists and returns a valid XmlElement
      pPrElement = CT_PPr.create();
      // <w:pPr> must be the first child if present
      element.children.insert(0, pPrElement);
    }
    return CT_PPr(pPrElement);
  }

  // --- Add Methods ---

  /// Add a new `<w:r>` child element and return its wrapper `CT_R`.
  /// The run is appended to the end of the paragraph's content.
  CT_R addR() {
    // <w:r> elements are appended after existing content (pPr, runs, hyperlinks etc.)
    // Assumes CT_R.create() exists and returns a valid XmlElement
    final rElement = CT_R.create();
    element.children.add(rElement);
    return CT_R(rElement);
  }

  /// Alias for `addR` to match Python hint `add_r`.
  CT_R add_r() => addR();

  /// Return a new `<w:p>` element inserted directly prior to this one in the
  /// parent element's children list.
  /// Throws [StateError] if this element has no parent or is not found in its parent.
  CT_P addPBefore() {
    final newPElement = OxmlElement(qnTagName); // Create a new <w:p>
    final parent = element.parent;
    if (parent == null || parent is! XmlElement) {
      throw StateError(
          'Cannot add paragraph before element with no parent element.');
    }
    final index = parent.children.indexOf(element);
    if (index == -1) {
      throw StateError(
          'Cannot add paragraph before element not found in parent.');
    }
    parent.children.insert(index, newPElement);
    return CT_P(newPElement);
  }

  /// Alias for `addPBefore` to match Python's `add_p_before`.
  CT_P add_p_before() => addPBefore();

  // --- Properties ---

  /// The paragraph alignment value (e.g., `WD_PARAGRAPH_ALIGNMENT.center`),
  /// derived from the `<w:jc>` grandchild element. Returns `null` if no
  /// alignment is explicitly set (inherits from style).
  WD_PARAGRAPH_ALIGNMENT? get alignment => pPr?.jcVal;

  /// Set the paragraph alignment. Adds `<w:pPr>` and `<w:jc>` child
  /// elements if necessary. Setting the value to `null` removes the
  /// explicit alignment setting, causing inheritance from the style hierarchy.
  set alignment(WD_PARAGRAPH_ALIGNMENT? value) {
    getOrAddPPr().jcVal = value;
  }

  /// Remove all child elements representing paragraph content (like `<w:r>`
  /// and `<w:hyperlink>`), leaving only the `<w:pPr>` (Paragraph Properties)
  /// element if it exists.
  void clearContent() {
    final pPrElement = childOrNull(CT_PPr.qnTagName);
    final childrenToRemove = element.children
        .where((node) => node is XmlElement && node != pPrElement)
        .toList();
    for (final child in childrenToRemove) {
      element.children.remove(child);
    }
  }

  /// Returns a list of the run (`CT_R`) and hyperlink (`CT_Hyperlink`)
  /// child elements of this paragraph, in the order they appear in the XML.
  List<BaseOxmlElement> get innerContentElements {
    final content = <BaseOxmlElement>[];
    final wUri = nsmap['w']!;
    for (final node in element.children) {
      if (node is XmlElement) {
        if (node.name.local == 'r' && node.name.namespaceUri == wUri) {
          content.add(CT_R(node));
        } else if (node.name.local == 'hyperlink' &&
            node.name.namespaceUri == wUri) {
          content.add(CT_Hyperlink(node));
        }
        // Future: Add checks for other potential inline content types
      }
    }
    return content;
  }

  /// Returns a list of all `w:lastRenderedPageBreak` elements found within
  /// the runs (`<w:r>`) or hyperlinks (`<w:hyperlink>`) of this paragraph.
  List<CT_LastRenderedPageBreak> get lastRenderedPageBreaks {
    final breaks = <CT_LastRenderedPageBreak>[];
    for (final contentElement in innerContentElements) {
      if (contentElement is CT_R) {
        breaks.addAll(contentElement.lastRenderedPageBreaks);
      } else if (contentElement is CT_Hyperlink) {
        breaks.addAll(contentElement.lastRenderedPageBreaks);
      }
    }
    return breaks;
  }

  /// Replace or add the `<w:sectPr>` grandchild element within the `<w:pPr>`
  /// child, ensuring it's placed correctly according to the OOXML sequence.
  /// [sectPr] should be the `CT_SectPr` object to insert.
  void setSectPr(CT_SectPr sectPr) {
    final pPr = getOrAddPPr();
    // Assuming CT_PPr has these methods implemented:
    pPr.removeSectPr();
    pPr.insertSectPr(sectPr);
  }

  /// The style ID string from the `w:val` attribute of the
  /// `./w:pPr/w:pStyle` grandchild element. Returns `null` if the paragraph
  /// has no explicit style applied (inherits default).
  String? get style => pPr?.style;

  /// Set the paragraph style by its ID. Adds `<w:pPr>` and `<w:pStyle>`
  /// child elements if necessary. Setting the value to `null` removes the
  /// explicit style setting, causing it to inherit the default paragraph style.
  set style(String? styleId) {
    getOrAddPPr().style = styleId;
  }

  /// Gets the concatenated text content of all runs (`<w:r>`) and hyperlinks
  /// (`<w:hyperlink>`) within this paragraph.
  /// Special characters like tabs (`<w:tab>`) and breaks (`<w:br>`, `<w:cr>`)
  /// within runs are converted to their string representations (`\t`, `\n`).
  String get text {
    final buffer = StringBuffer();
    for (final contentElement in innerContentElements) {
      // --- CORRECTION: Check type before accessing .text ---
      if (contentElement is CT_R) {
        // Assuming CT_R has a 'text' getter that handles its children (<w:t> etc.)
        buffer.write(contentElement.text);
      } else if (contentElement is CT_Hyperlink) {
        // Assuming CT_Hyperlink has a 'text' getter that handles its runs
        buffer.write(contentElement.text);
      }
      // --- Add else if blocks here for other text-contributing element types ---
      // else if (contentElement is CT_SimpleField) { ... }
    }
    return buffer.toString();
  }

  /// Qualified name for the `<w:p>` element -> "{namespace}p".
  static final qnTagName = qn('w:p');

  /// Static factory method to create a new minimal `<w:p>` [XmlElement].
  static XmlElement create() => OxmlElement(qnTagName);
}

/// Path: lib/src/oxml/document.dart
/// Based on python-docx: docx/oxml/document.py
/// Custom element classes that correspond to the document part (<w:document>, <w:body>).

import 'package:xml/xml.dart';
import 'ns.dart' show qn; // qn function
import 'parser.dart' show OxmlElement; // OxmlElement factory
import 'section.dart' show CT_SectPr;
import 'table.dart' show CT_Tbl;
import 'text/paragraph.dart' show CT_P;
import 'xmlchemy.dart' show BaseOxmlElement;
import 'xmlchemy_descriptors.dart'; // Descriptors like ZeroOrOne, ZeroOrMore

/// `<w:body>`, the container element for the main document story in `document.xml`.
class CT_Body extends BaseOxmlElement {
  CT_Body(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:body');

  // --- Define sequence for insertion ---
  // Paragraphs and Tables can intersperse, both come before the final sectPr.
  static final childSequence = [qn('w:p'), qn('w:tbl'), qn('w:sectPr')];

  // --- Descriptors ---
  static final _p = ZeroOrMore<CT_P>(qn('w:p'), (el) => CT_P(el));
  static final _tbl = ZeroOrMore<CT_Tbl>(qn('w:tbl'), (el) => CT_Tbl(el));
  // sectPr is the last element
  static final _sectPr = ZeroOrOne<CT_SectPr>(qn('w:sectPr'), successors: []);

  // --- Property Accessors ---
  List<CT_P> get pList => _p.getElements(this);

  /// Alias for pList
  List<CT_P> get p_lst => pList;

  List<CT_Tbl> get tblList => _tbl.getElements(this);

  /// Alias for tblList
  List<CT_Tbl> get tbl_lst => tblList;

  CT_SectPr? get sectPr => _sectPr.getElement(this, (el) => CT_SectPr(el));
  CT_SectPr getOrAddSectPr() =>
      _sectPr.getOrAdd(this, CT_SectPr.create, (el) => CT_SectPr(el));

  /// Alias for getOrAddSectPr
  CT_SectPr get_or_add_sectPr() => getOrAddSectPr();

  // --- Methods ---

  /// Adds a new `<w:p>` element before the final `<w:sectPr>` if present,
  /// otherwise appends it. Returns the new paragraph wrapper.
  CT_P addP() {
    final pElement = CT_P.create();
    // Insert before the final sectPr
    insertChild(pElement, [qn('w:sectPr')]);
    return CT_P(pElement);
  }

  /// Alias for addP
  CT_P add_p() => addP();

  /// Inserts a `<w:tbl>` element before the final `<w:sectPr>` if present,
  /// otherwise appends it.
  void insertTbl(CT_Tbl tbl) {
    // Detach tbl if it's already parented elsewhere
    if (tbl.element.parent != null) {
      tbl.element.parent!.children.remove(tbl.element);
    }
    // Insert before the final sectPr
    insertChild(tbl.element, [qn('w:sectPr')]);
  }

  /// Alias for insertTbl
  void insert_tbl(CT_Tbl tbl) => insertTbl(tbl);

  /// Adds a section break at the end of the document body.
  /// Returns the *new* sentinel `<w:sectPr>` (the one now at the very end).
  /// This involves cloning the existing last section's properties into a new
  /// paragraph before the end, and modifying the original sentinel.
  CT_SectPr addSectionBreak() {
    // 1. Get or add the current last <w:sectPr> (the sentinel)
    final sentinelSectPr = getOrAddSectPr();

    // 2. Clone it *before* modifying it
    // Assumes CT_SectPr.clone() exists and performs a deep copy, removing rsids
    final clonedSectPr = sentinelSectPr.clone();

    // 3. Add a new paragraph *before* the sentinel sectPr and set its sectPr
    //    to the cloned properties. This new paragraph marks the end of the *previous* section.
    final newPara = addP(); // Adds paragraph before sentinel sectPr
    newPara.setSectPr(clonedSectPr); // Assumes CT_P.setSectPr exists

    // 4. Remove header/footer references from the *original* sentinel <w:sectPr>.
    //    This makes the *new* last section inherit from the previous one.
    // Assumes CT_SectPr has removeHeaderReference and removeFooterReference methods
    // Need to iterate over *copies* of the lists while removing
    final headerRefsToRemove = sentinelSectPr.headerReference.toList();
    for (final hdrRef in headerRefsToRemove) {
      sentinelSectPr.removeHeaderReference(hdrRef.type);
    }
    final footerRefsToRemove = sentinelSectPr.footerReference.toList();
    for (final ftrRef in footerRefsToRemove) {
      sentinelSectPr.removeFooterReference(ftrRef.type);
    }

    // 5. The original sentinelSectPr (now modified) defines the new last section.
    return sentinelSectPr;
  }

  /// Alias for addSectionBreak
  CT_SectPr add_section_break() => addSectionBreak();

  /// Removes all content child elements (`<w:p>`, `<w:tbl>`, etc.) from this
  /// `<w:body>` element, leaving the final `<w:sectPr>` element if present.
  void clearContent() {
    final sectPrElement = sectPr?.element; // Get the element if it exists
    // Collect elements to remove, excluding the sectPr
    final childrenToRemove = element.children
        .where((node) => node is XmlElement && node != sectPrElement)
        .toList();
    for (final child in childrenToRemove) {
      element.children.remove(child);
    }
  }

  /// Alias for clearContent
  void clear_content() => clearContent();

  /// Returns child `<w:p>` and `<w:tbl>` elements in document order.
  /// This provides a simple view, ignoring potential wrappers like <w:ins>.
  List<BaseOxmlElement /* CT_P | CT_Tbl */ > get innerContentElements {
    final content = <BaseOxmlElement>[];
    final pTag = CT_P.qnTagName;
    final tblTag = CT_Tbl.qnTagName;
    for (final child in element.children.whereType<XmlElement>()) {
      // Check only for p and tbl tags directly under body
      if (child.name.qualified == pTag) {
        content.add(CT_P(child));
      } else if (child.name.qualified == tblTag) {
        content.add(CT_Tbl(child));
      }
    }
    return content;
  }

  /// Alias for innerContentElements getter
  List<BaseOxmlElement /* CT_P | CT_Tbl */ > get inner_content_elements =>
      innerContentElements;
}

/// `<w:document>` element, the root element of a document.xml file.
class CT_Document extends BaseOxmlElement {
  CT_Document(super.element);

  static XmlElement create() {
    // --- CORRECTION: Create parent, then add child ---
    final doc = OxmlElement(qnTagName);
    doc.children.add(CT_Body.create()); // Add required body
    return doc;
    // --- End Correction ---
  }


  static final qnTagName = qn('w:document');

  // --- Descriptors ---
  static final _body =
      OneAndOnlyOne<CT_Body>(CT_Body.qnTagName, (el) => CT_Body(el));

  // --- Property Accessors ---
  /// The required `<w:body>` child element.
  CT_Body get body => _body.getElement(this);

  /// Read-only list of all `<w:sectPr>` elements in the document body,
  /// in document order. Includes the final one in `<w:body>` and any
  /// preceding ones within `<w:p>/<w:pPr>`.
  List<CT_SectPr> get sectPrList {
    final sectPrs = <CT_SectPr>[];
    final bodyElm = body;

    for (final paragraph in bodyElm.pList) {
      final pSectPr = paragraph.pPr?.sectPr;
      if (pSectPr != null) {
        sectPrs.add(pSectPr);
      }
    }

    final trailingSectPr = bodyElm.sectPr;
    if (trailingSectPr != null) {
      sectPrs.add(trailingSectPr);
    }

    if (sectPrs.isEmpty) {
      sectPrs.add(bodyElm.getOrAddSectPr());
    }

    return sectPrs;
  }

  /// Alias for sectPrList
  List<CT_SectPr> get sectPr_lst => sectPrList;
}

// Assume necessary Converters and BaseSimpleType are defined/imported
// Assume Enums have fromXml/xmlValue capabilities

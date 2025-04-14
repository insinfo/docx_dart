/// Path: lib/src/oxml/text/pagebreak.dart
/// Based on python-docx: docx/oxml/text/pagebreak.py
/// Custom element class for rendered page-break (CT_LastRenderedPageBreak).

import 'package:xml/xml.dart'; // Importa XmlElement, XmlNode e outros

// Importa qn e nsmap
import '../ns.dart' show nsmap, qn;
import '../parser.dart' show OxmlElement; // OxmlElement para create
import '../xmlchemy.dart' show BaseOxmlElement; // BaseOxmlElement
import 'hyperlink.dart' show CT_Hyperlink; // CT_Hyperlink
import 'paragraph.dart' show CT_P; // CT_P
import 'run.dart' show CT_R; // CT_R

// ignore_for_file: camel_case_types, unnecessary_this

/// `<w:lastRenderedPageBreak>` element, indicating page break inserted by renderer.
// ... (Doc comments remain the same) ...
class CT_LastRenderedPageBreak extends BaseOxmlElement {
  CT_LastRenderedPageBreak(super.element);

  /// Creates a new `XmlElement` suitable for a CT_LastRenderedPageBreak.
  static XmlElement create() => OxmlElement(qnTagName);

  static final qnTagName = qn('w:lastRenderedPageBreak');

  /// A "loose" [CT_P] containing only the paragraph content *following* this break.
  // ... (Doc comments remain the same) ...
  CT_P get followingFragmentP {
    final enclosingP = this._enclosingP;
    final firstLrpbElement = _firstLrpbInP(enclosingP)?.element;
    if (enclosingP == null ||
        firstLrpbElement == null ||
        this.element != firstLrpbElement) {
      throw ArgumentError(
          "followingFragmentP only defined on first rendered page-break in paragraph");
    }
    return _isInHyperlink ? _followingFragInHlink : _followingFragInRun;
  }

  /// `true` when this page-break element is the last "content" in the paragraph.
  // ... (Doc comments remain the same) ...
  bool get followsAllContent {
    // ... (implementation remains the same) ...
    if (_isInHyperlink) return false;
    final p = _enclosingP;
    if (p == null) return false;
    final runs = p.rElements;
    if (runs.isEmpty) return false;

    final lastRunElement = runs.last.element;
    final parentRunElement = this.element.parent;

    if (parentRunElement != lastRunElement) {
      return false;
    }

    final lastRunChildren = lastRunElement.children;
    final lrpbIndex = lastRunChildren.indexOf(this.element);
    if (lrpbIndex == -1) return false;

    final wUri = nsmap['w']!;
    for (int i = lrpbIndex + 1; i < lastRunChildren.length; i++) {
      final sibling = lastRunChildren[i];
      if (sibling is XmlElement &&
          _runInnerContentTypes.contains(sibling.name.local) &&
          sibling.name.namespaceUri == wUri) {
        return false;
      }
    }
    return true;
  }

  /// `true` when a `w:lastRenderedPageBreak` precedes all paragraph content.
  // ... (Doc comments remain the same) ...
  bool get precedesAllContent {
    // ... (implementation remains the same) ...
    if (_isInHyperlink) return false;
    final p = _enclosingP;
    if (p == null) return false;
    final runs = p.rElements;
    if (runs.isEmpty) return false;

    final firstRunElement = runs.first.element;
    final parentRunElement = this.element.parent;

    if (parentRunElement != firstRunElement) {
      return false;
    }

    final firstRunChildren = firstRunElement.children;
    final lrpbIndex = firstRunChildren.indexOf(this.element);
    if (lrpbIndex == -1) return false;

    final wUri = nsmap['w']!;
    for (int i = 0; i < lrpbIndex; i++) {
      final sibling = firstRunChildren[i];
      if (sibling is XmlElement && sibling.name.qualified == qn('w:rPr')) {
        continue;
      }
      if (sibling is XmlElement &&
          _runInnerContentTypes.contains(sibling.name.local) &&
          sibling.name.namespaceUri == wUri) {
        return false;
      }
    }
    return true;
  }

  /// A "loose" [CT_P] containing only the paragraph content *before* this break.
  // ... (Doc comments remain the same) ...
  CT_P get precedingFragmentP {
    final enclosingP = this._enclosingP;
    final firstLrpbElement = _firstLrpbInP(enclosingP)?.element;
    if (enclosingP == null ||
        firstLrpbElement == null ||
        this.element != firstLrpbElement) {
      throw ArgumentError(
          "precedingFragmentP only defined on first rendered page-break in paragraph");
    }
    return _isInHyperlink ? _precedingFragInHlink : _precedingFragInRun;
  }

  // --- Private Properties and Methods ---

  CT_Hyperlink _enclosingHyperlink(CT_LastRenderedPageBreak lrpb) {
    // ... (implementation remains the same) ...
    final rParent = lrpb.element.parent;
    final hyperlinkParent = rParent?.parent;
    final wUri = nsmap['w']!;
    if (rParent == null ||
        !(rParent is XmlElement) ||
        rParent.name.local != 'r' ||
        rParent.name.namespaceUri != wUri) {
      throw StateError('CT_LastRenderedPageBreak parent is not w:r');
    }
    if (hyperlinkParent != null &&
        hyperlinkParent is XmlElement &&
        hyperlinkParent.name.local == 'hyperlink' &&
        hyperlinkParent.name.namespaceUri == wUri) {
      return CT_Hyperlink(hyperlinkParent);
    }
    throw StateError('CT_LastRenderedPageBreak is not inside a w:hyperlink');
  }

  CT_P? get _enclosingP {
    // ... (implementation remains the same) ...
    final wUri = nsmap['w']!;
    XmlNode? current = element;
    while (current != null) {
      if (current is XmlElement &&
          current.name.local == 'p' &&
          current.name.namespaceUri == wUri) {
        return CT_P(current);
      }
      current = current.parent;
    }
    return null;
  }

  CT_LastRenderedPageBreak? _firstLrpbInP(CT_P? p) {
    // ... (implementation remains the same) ...
    if (p == null) return null;
    final lrpbs = <CT_LastRenderedPageBreak>[];
    for (final r in p.rElements) {
      lrpbs.addAll(r.childrenWhereType<CT_LastRenderedPageBreak>(
          CT_LastRenderedPageBreak.qnTagName,
          (el) => CT_LastRenderedPageBreak(el)));
    }
    for (final h in p.hyperlinkElements) {
      for (final r in h.rElements) {
        lrpbs.addAll(r.childrenWhereType<CT_LastRenderedPageBreak>(
            CT_LastRenderedPageBreak.qnTagName,
            (el) => CT_LastRenderedPageBreak(el)));
      }
    }
    return lrpbs.isEmpty ? null : lrpbs.first;
  }

  /// Following [CT_P] fragment when break occurs within a hyperlink.
  CT_P get _followingFragInHlink {
    final pElement = _enclosingP?.element;
    if (pElement == null)
      throw StateError("Cannot get fragment, enclosing paragraph not found.");

    // --- CORRECTION: Use 'as XmlElement' ---
    final pElementClone = pElement.copy();
    // --- End Correction ---
    final p = CT_P(pElementClone);

    final lrpbInClone = _firstLrpbInP(p);
    if (lrpbInClone == null)
      throw StateError("LRPB not found in cloned paragraph.");
    final hyperlinkInClone = _enclosingHyperlink(lrpbInClone);
    final hyperlinkElement = hyperlinkInClone.element;

    final parentChildren = p.element.children;
    final hyperlinkIndex = parentChildren.indexOf(hyperlinkElement);
    if (hyperlinkIndex == -1)
      throw StateError("Hyperlink not found in cloned paragraph children.");

    // Remove elements *before* the hyperlink, preserving pPr
    final List<XmlNode> toRemove = [];
    for (int i = 0; i < hyperlinkIndex; i++) {
      final child = parentChildren[i];
      if (!(child is XmlElement && child.name.qualified == qn('w:pPr'))) {
        toRemove.add(child);
      }
    }
    for (final node in toRemove) {
      parentChildren.remove(node);
    }

    // Remove the whole hyperlink
    parentChildren.remove(hyperlinkElement);

    return p;
  }

  /// Following [CT_P] fragment when break does not occur in a hyperlink.
  CT_P get _followingFragInRun {
    final pElement = _enclosingP?.element;
    if (pElement == null)
      throw StateError("Cannot get fragment, enclosing paragraph not found.");

    // --- CORRECTION: Use 'as XmlElement' ---
    final pElementClone = pElement.copy();
    // --- End Correction ---
    final p = CT_P(pElementClone);

    final lrpbInClone = _firstLrpbInP(p);
    if (lrpbInClone == null)
      throw StateError("LRPB not found in cloned paragraph.");

    final enclosingRElement = lrpbInClone.element.parent;
    final wUri = nsmap['w']!;
    if (enclosingRElement == null ||
        !(enclosingRElement is XmlElement) ||
        enclosingRElement.name.local != 'r' ||
        enclosingRElement.name.namespaceUri != wUri) {
      throw StateError('CT_LastRenderedPageBreak must be child of w:r');
    }
    final enclosingR = CT_R(enclosingRElement);

    // Remove preceding in P using index
    final pChildren = p.element.children;
    final rIndexInP = pChildren.indexOf(enclosingRElement);
    if (rIndexInP == -1)
      throw StateError("Enclosing run not found in cloned paragraph children.");
    final pToRemove = <XmlNode>[];
    for (int i = 0; i < rIndexInP; i++) {
      final child = pChildren[i];
      if (!(child is XmlElement && child.name.qualified == qn('w:pPr'))) {
        pToRemove.add(child);
      }
    }
    for (final node in pToRemove) {
      pChildren.remove(node);
    }

    // Remove preceding in R using index
    final rChildren = enclosingR.element.children;
    final lrpbIndexInR = rChildren.indexOf(lrpbInClone.element);
    if (lrpbIndexInR == -1)
      throw StateError("LRPB not found in cloned run children.");
    final rToRemove = <XmlNode>[];
    for (int i = 0; i < lrpbIndexInR; i++) {
      final child = rChildren[i];
      if (!(child is XmlElement && child.name.qualified == qn('w:rPr'))) {
        rToRemove.add(child);
      }
    }
    for (final node in rToRemove) {
      rChildren.remove(node);
    }

    // Remove the page-break itself
    enclosingR.element.children.remove(lrpbInClone.element);

    return p;
  }

  /// `true` when this page-break is embedded in a hyperlink run.
  bool get _isInHyperlink {
    // ... (implementation remains the same) ...
    final rParent = element.parent;
    final hyperlinkParent = rParent?.parent;
    final wUri = nsmap['w']!;
    return hyperlinkParent != null &&
        hyperlinkParent is XmlElement &&
        hyperlinkParent.name.local == 'hyperlink' &&
        hyperlinkParent.name.namespaceUri == wUri;
  }

  /// Preceding [CT_P] fragment when break occurs within a hyperlink.
  CT_P get _precedingFragInHlink {
    final pElement = _enclosingP?.element;
    if (pElement == null)
      throw StateError("Cannot get fragment, enclosing paragraph not found.");

    // --- CORRECTION: Use 'as XmlElement' ---
    final pElementClone = pElement.copy();
    // --- End Correction ---
    final p = CT_P(pElementClone);

    final lrpbInClone = _firstLrpbInP(p);
    if (lrpbInClone == null)
      throw StateError("LRPB not found in cloned paragraph.");
    final hyperlinkInClone = _enclosingHyperlink(lrpbInClone);
    final hyperlinkElement = hyperlinkInClone.element;

    // Remove following using index
    final parentChildren = p.element.children;
    final hyperlinkIndex = parentChildren.indexOf(hyperlinkElement);
    if (hyperlinkIndex == -1)
      throw StateError("Hyperlink not found in cloned paragraph children.");
    final toRemove = parentChildren.sublist(hyperlinkIndex + 1).toList();
    for (final e in toRemove) {
      parentChildren.remove(e);
    }

    // Remove this page-break from inside the hyperlink
    final lrpbParent = lrpbInClone.element.parent;
    if (lrpbParent == null) throw StateError("Cloned LRPB has no parent.");
    lrpbParent.children.remove(lrpbInClone.element);

    return p;
  }

  /// Preceding [CT_P] fragment when break does not occur in a hyperlink.
  CT_P get _precedingFragInRun {
    final pElement = _enclosingP?.element;
    if (pElement == null)
      throw StateError("Cannot get fragment, enclosing paragraph not found.");

    // --- CORRECTION: Use 'as XmlElement' ---
    final pElementClone = pElement.copy();
    // --- End Correction ---
    final p = CT_P(pElementClone);

    final lrpbInClone = _firstLrpbInP(p);
    if (lrpbInClone == null)
      throw StateError("LRPB not found in cloned paragraph.");

    final enclosingRElement = lrpbInClone.element.parent;
    final wUri = nsmap['w']!;
    if (enclosingRElement == null ||
        !(enclosingRElement is XmlElement) ||
        enclosingRElement.name.local != 'r' ||
        enclosingRElement.name.namespaceUri != wUri) {
      throw StateError('CT_LastRenderedPageBreak must be child of w:r');
    }
    final enclosingR = CT_R(enclosingRElement);

    // Remove following in P using index
    final pChildren = p.element.children;
    final rIndexInP = pChildren.indexOf(enclosingRElement);
    if (rIndexInP == -1)
      throw StateError("Enclosing run not found in cloned paragraph children.");
    final pToRemove = pChildren.sublist(rIndexInP + 1).toList();
    for (final e in pToRemove) {
      pChildren.remove(e);
    }

    // Remove following in R using index
    final rChildren = enclosingR.element.children;
    final lrpbIndexInR = rChildren.indexOf(lrpbInClone.element);
    if (lrpbIndexInR == -1)
      throw StateError("LRPB not found in cloned run children.");
    final rToRemove = rChildren.sublist(lrpbIndexInR + 1).toList();
    for (final e in rToRemove) {
      rChildren.remove(e);
    }

    // Remove the page-break itself
    enclosingR.element.children.remove(lrpbInClone.element);

    return p;
  }

  /// Set of local names for run inner-content elements considered "content".
  static const Set<String> _runInnerContentTypes = {
    'br',
    'cr',
    'drawing',
    'noBreakHyphen',
    'ptab',
    't',
    'tab',
  };
}

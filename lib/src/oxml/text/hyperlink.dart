/// Path: lib/src/oxml/text/hyperlink.dart
/// Based on python-docx: docx/oxml/text/hyperlink.py
/// Custom element classes related to hyperlinks (CT_Hyperlink).

import '../ns.dart'; // qn
import '../simpletypes.dart'; // ST_OnOff, ST_String, XsdString converters
import '../xmlchemy.dart'; // BaseOxmlElement, OptionalAttribute, ZeroOrMore
import 'run.dart'; // CT_R
import 'pagebreak.dart'; // CT_LastRenderedPageBreak

/// `<w:hyperlink>` element, containing the text and address for a hyperlink.
class CT_Hyperlink extends BaseOxmlElement {
  CT_Hyperlink(super.element);

  /// Relationship Id linking to the external target URL or internal bookmark.
  String? get rId => getAttrVal('r:id', xsdStringConverter);
  set rId(String? value) => setAttrVal('r:id', value, xsdStringConverter);

  /// Target frame or bookmark within the document.
  String? get anchor => getAttrVal('w:anchor', stStringConverter);
  set anchor(String? value) => setAttrVal('w:anchor', value, stStringConverter);

  /// Specifies whether this hyperlink should be added to the history. Defaults to true.
  bool get history =>
      // Use ?? para fornecer o valor padrão CASO getAttrVal retorne null
      getAttrVal<bool>('w:history', stOnOffConverter, defaultValue: true) ??
      true;

  set history(bool value) =>
      setAttrVal<bool>('w:history', value, stOnOffConverter,
          defaultValue: true);

  // --- Child Element Access ---

  /// Sequence of `<w:r>` elements contained in this hyperlink.
  List<CT_R> get rElements => childrenWhereType<CT_R>(
        CT_R.qnTagName,
        (el) => CT_R(el),
      );

  /// Alias for `rElements` to match Python's `r_lst`.
  List<CT_R> get r_lst => rElements;

  // --- Properties ---

  /// All `w:lastRenderedPageBreak` descendants of this hyperlink.
  List<CT_LastRenderedPageBreak> get lastRenderedPageBreaks {
    final breaks = <CT_LastRenderedPageBreak>[];
    for (final r in rElements) {
      // Assumes CT_R has a 'lastRenderedPageBreaks' getter that returns List<CT_LastRenderedPageBreak>
      breaks.addAll(r.lastRenderedPageBreaks);
    }
    return breaks;
  }

  /// The textual content of this hyperlink.
  /// `CT_Hyperlink` stores the hyperlink-text as one or more `w:r` children.
  String get text {
    final buffer = StringBuffer();
    for (final r in rElements) {
      // Assumes CT_R has a 'text' getter that returns its textual content
      buffer.write(r.text);
    }
    return buffer.toString();
  }

  /// Qualified name for the <w:hyperlink> element.
  static final qnTagName = qn('w:hyperlink');
}

// --- Placeholder Converters (replace with actual implementation) ---
final xsdStringConverter = XsdStringConverter();
final stStringConverter = ST_StringConverter();
final stOnOffConverter = ST_OnOffConverter();

class XsdStringConverter extends BaseSimpleType<String> {
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;

  @override
  void validate(String value) {}
}

class ST_StringConverter extends BaseSimpleType<String> {
  @override
  String fromXml(String xmlValue) => xmlValue;
  @override
  String? toXml(String? value) => value;

  @override
  void validate(String value) {}
}

class ST_OnOffConverter extends BaseSimpleType<bool> {
  @override
  bool fromXml(String xmlValue) =>
      xmlValue == '1' || xmlValue == 'true' || xmlValue == 'on';
  @override
  String? toXml(bool? value) => value == null ? null : (value ? '1' : '0');

  @override
  void validate(bool value) {}
}

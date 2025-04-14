/// Path: lib/src/oxml/numbering.dart
/// Based on python-docx: docx/oxml/numbering.py
/// Custom element classes related to the numbering part (<w:numbering>, <w:num>, etc.).

import 'package:docx_dart/src/oxml/shared.dart';
import 'package:xml/xml.dart';
import 'package:collection/collection.dart'; // For firstWhereOrNull

import 'ns.dart' show qn; // qn function
import 'parser.dart' show OxmlElement; // OxmlElement factory

import 'simpletypes.dart';
import 'xmlchemy.dart' show BaseOxmlElement;

import 'xmlchemy_descriptors.dart';

/// `<w:num>` element, representing a concrete list definition instance.
/// References an abstract numbering definition via `<w:abstractNumId>`.
class CT_Num extends BaseOxmlElement {
  CT_Num(super.element);

  /// Creates a new `<w:num>` element with the specified numId and abstractNumId.
  static XmlElement create({required int numId, required int abstractNumId}) {
    final numElement =
        OxmlElement(qnTagName, attrs: {qn('w:numId'): numId.toString()});
    // Create and append the required abstractNumId child
    numElement.children.add(
        CT_DecimalNumber.create(CT_Num._abstractNumIdTagName, abstractNumId));
    return numElement;
  }

  static final qnTagName = qn('w:num');
  // Define qualified names for children used in descriptors
  static final _abstractNumIdTagName = qn('w:abstractNumId');
  static final _lvlOverrideTagName = qn('w:lvlOverride');

  // --- Descriptors ---
  static final _abstractNumId = OneAndOnlyOne<CT_DecimalNumber>(
      _abstractNumIdTagName, (el) => CT_DecimalNumber(el));
  static final _lvlOverride =
      ZeroOrMore<CT_NumLvl>(_lvlOverrideTagName, (el) => CT_NumLvl(el));

  // --- Attributes ---
  /// The unique ID for this concrete numbering definition. Required attribute.
  int get numId => getReqAttrVal('w:numId', stDecimalNumberConverter);
  set numId(int value) =>
      setReqAttrVal('w:numId', value, stDecimalNumberConverter);

  // --- Child Element Access ---
  /// The required `<w:abstractNumId>` child element.
  CT_DecimalNumber get abstractNumId => _abstractNumId.getElement(this);

  /// List of `<w:lvlOverride>` child elements.
  List<CT_NumLvl> get lvlOverride => _lvlOverride.getElements(this);

  /// Alias for lvlOverride
  List<CT_NumLvl> get lvlOverride_lst => lvlOverride;

  // --- Methods ---

  /// Adds a new `<w:lvlOverride>` child element for the specified level [ilvl]
  /// and returns its wrapper.
  CT_NumLvl addLvlOverride(int ilvl) {
    // Create the new lvlOverride element with the required ilvl attribute
    final lvlOverrideElement = CT_NumLvl.create(ilvl: ilvl);
    // Append the new element (order might matter based on schema, appending is simplest)
    // Use addChildElement helper if sequence matters more strictly.
    element.children.add(lvlOverrideElement);
    return CT_NumLvl(lvlOverrideElement);
  }

  /// Alias for addLvlOverride
  CT_NumLvl add_lvlOverride(int ilvl) => addLvlOverride(ilvl);

  /// Static factory method provided for compatibility/clarity, uses create().
  static CT_Num newNum(int numId, int abstractNumId) {
    return CT_Num(create(numId: numId, abstractNumId: abstractNumId));
  }
}

/// `<w:lvlOverride>` element, overrides formatting for a specific level
/// within a concrete numbering definition (<w:num>).
class CT_NumLvl extends BaseOxmlElement {
  CT_NumLvl(super.element);

  /// Creates a new `<w:lvlOverride>` element with the required `ilvl` attribute.
  static XmlElement create({required int ilvl}) =>
      OxmlElement(qnTagName, attrs: {qn('w:ilvl'): ilvl.toString()});
  static final qnTagName = qn('w:lvlOverride');
  // Define qualified name for child used in descriptor
  static final _startOverrideTagName = qn('w:startOverride');

  // --- Descriptors ---
  // Successors likely include <w:lvl> if implemented
  static final _startOverride = ZeroOrOne<CT_DecimalNumber>(
      _startOverrideTagName,
      successors: [qn('w:lvl')]);

  // --- Attributes ---
  /// The level being overridden (0-based). Required attribute.
  int get ilvl => getReqAttrVal('w:ilvl', stDecimalNumberConverter);
  set ilvl(int value) =>
      setReqAttrVal('w:ilvl', value, stDecimalNumberConverter);

  // --- Child Element Access ---
  /// The optional `<w:startOverride>` child element.
  CT_DecimalNumber? get startOverride =>
      _startOverride.getElement(this, (el) => CT_DecimalNumber(el));

  // --- Methods ---

  /// Adds a new `<w:startOverride>` child element with the specified value [val]
  /// and returns its wrapper.
  CT_DecimalNumber addStartOverride(int val) {
    // Use the descriptor's getOrAdd method
    // Pass the factory for CT_DecimalNumber which takes tag name and value
    final startOverride = _startOverride.getOrAdd(
        this,
        () => CT_DecimalNumber.create(_startOverrideTagName, val),
        (el) => CT_DecimalNumber(el));
    // Ensure the value is set (getOrAdd might return existing with different val)
    startOverride.val = val;
    return startOverride;
  }

  /// Alias for addStartOverride
  CT_DecimalNumber add_startOverride(int val) => addStartOverride(val);
}

/// `<w:numPr>` element, container for numbering properties applied to a paragraph.
class CT_NumPr extends BaseOxmlElement {
  CT_NumPr(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:numPr');
  // Define qualified names for children
  static final _ilvlTagName = qn('w:ilvl');
  static final _numIdTagName = qn('w:numId');

  // --- Define sequence ---
  static final _childSequence = [
    _ilvlTagName,
    _numIdTagName,
    qn('w:numberingChange'),
    qn('w:ins')
  ];

  // --- Descriptors ---
  static final _ilvl = ZeroOrOne<CT_DecimalNumber>(_ilvlTagName,
      successors: _childSequence.sublist(1));
  static final _numId = ZeroOrOne<CT_DecimalNumber>(_numIdTagName,
      successors: _childSequence.sublist(2));

  // --- Child Element Access ---
  /// Optional `<w:ilvl>` child element specifying the numbering level.
  CT_DecimalNumber? get ilvl =>
      _ilvl.getElement(this, (el) => CT_DecimalNumber(el));

  /// Optional `<w:numId>` child element referencing a concrete numbering definition.
  CT_DecimalNumber? get numId =>
      _numId.getElement(this, (el) => CT_DecimalNumber(el));

  // --- Convenience Property Setters ---
  /// Sets the value of the `<w:ilvl>` child element. Adds/removes element as needed.
  set ilvlVal(int? value) {
    if (value == null) {
      _ilvl.remove(this);
    } else {
      _ilvl
          .getOrAdd(this, () => CT_DecimalNumber.create(_ilvlTagName, value),
              (el) => CT_DecimalNumber(el))
          .val = value;
    }
  }

  /// Sets the value of the `<w:numId>` child element. Adds/removes element as needed.
  set numIdVal(int? value) {
    if (value == null) {
      _numId.remove(this);
    } else {
      _numId
          .getOrAdd(this, () => CT_DecimalNumber.create(_numIdTagName, value),
              (el) => CT_DecimalNumber(el))
          .val = value;
    }
  }
}

/// `<w:numbering>` element, the root element of a numbering part (numbering.xml).
class CT_Numbering extends BaseOxmlElement {
  CT_Numbering(super.element);
  static XmlElement create() => OxmlElement(qnTagName);
  static final qnTagName = qn('w:numbering');
  // Define sequence for <w:num> if needed (e.g., before <w:numIdMacAtCleanup>)
  static final _childSequence = [qn('w:numIdMacAtCleanup')];

  // --- Descriptors ---
  static final _num = ZeroOrMore<CT_Num>(CT_Num.qnTagName, (el) => CT_Num(el));

  // --- Child Element Access ---
  /// List of `<w:num>` child elements defining concrete numbering instances.
  List<CT_Num> get numList => _num.getElements(this);

  /// Alias for numList
  List<CT_Num> get num_lst => numList;

  // --- Methods ---

  /// Adds a new `<w:num>` element referencing abstract numbering definition
  /// [abstractNumId], using the next available `numId`. Returns the new wrapper.
  CT_Num addNum(int abstractNumId) {
    final nextId = this._nextNumId;
    final numElement =
        CT_Num.create(numId: nextId, abstractNumId: abstractNumId);
    // Use addChildElement helper for proper insertion if sequence matters
    addChildElement(_childSequence, () => numElement);
    return CT_Num(numElement);
  }

  /// Alias for addNum
  CT_Num add_num(int abstractNumId) => addNum(abstractNumId);

  /// Returns the `<w:num>` child element with matching `w:numId` attribute.
  /// Throws [KeyError] (simulated via ArgumentError) if not found.
  CT_Num numHavingNumId(int numId) {
    final numElement = numList.firstWhereOrNull((num) => num.numId == numId);
    if (numElement == null) {
      throw ArgumentError(
          "no <w:num> element with numId $numId"); // Simulating KeyError
    }
    return numElement;
  }

  /// Alias for numHavingNumId
  CT_Num num_having_numId(int numId) => numHavingNumId(numId);

  /// Calculates the next available `numId` for a new `<w:num>` element.
  /// Starts at 1 and fills gaps.
  int get _nextNumId {
    final numIds = numList.map((numEl) => numEl.numId).toList();
    numIds.sort(); // Sort to easily find the first gap
    int nextId = 1;
    for (final id in numIds) {
      if (id == nextId) {
        nextId++;
      } else if (id > nextId) {
        break; // Found a gap
      }
      // Skip if id < nextId (shouldn't happen with sort, but safe)
    }
    return nextId;
  }

  /// Alias for _nextNumId
  int get next_numId => _nextNumId;
}

// --- Placeholder Converters (Assume defined in simpletypes.dart) ---
final stDecimalNumberConverter = const ST_DecimalNumberConverter();

class ST_DecimalNumberConverter implements BaseSimpleType<int> {
  const ST_DecimalNumberConverter();
  @override
  int fromXml(String xmlValue) => int.parse(xmlValue);
  @override
  String? toXml(int? value) => value?.toString();
  @override
  void validate(int value) {} // Add range checks if needed
}

// Assume CT_DecimalNumber is defined in shared.dart and has a create(qn, val) method
// Assume BaseSimpleType exists and is imported
// Assume WD_SECTION_START, WD_HEADER_FOOTER, WD_ORIENTATION enums have fromXml/xmlValue

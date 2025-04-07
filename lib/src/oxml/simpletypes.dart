/// Path: lib/src/oxml/simpletypes.dart
/// Based on python-docx: docx/oxml/simpletypes.py
///
/// Simple-type classes corresponding to ST_* schema items.
/// These provide validation and format translation for values stored in XML element
/// attributes. Naming generally corresponds to the simple type in the associated XML
/// schema.



import 'package:docx_dart/src/enum/dml.dart';
import 'package:docx_dart/src/enum/text.dart';

import '../exceptions.dart'; // Assuming InvalidXmlError is defined here
import '../shared.dart'; // Assuming Length, Emu, Pt, Twips, RGBColor are here

/// Base interface for converting between Dart types and their XML string representations.
/// Also includes validation.
abstract class BaseSimpleType<T> {
  /// Converts an XML string value to the Dart type [T].
  /// Throws [InvalidXmlError] or [FormatException] if conversion fails.
  T fromXml(String xmlValue);

  /// Converts a Dart value of type [T] to its XML string representation.
  /// Performs validation before conversion.
  /// Returns null if the [value] corresponds to attribute removal (e.g., a default).
  /// Throws [ArgumentError] or [RangeError] if validation fails or conversion is not supported.
  String? toXml(T value);

  /// Validates the Dart value. Called by [toXml].
  /// Throws [ArgumentError] or [RangeError] if invalid.
  void validate(T value);
}

// --- Validation Helpers ---

void _validateInt(dynamic value) {
  if (value is! int) {
    throw ArgumentError("value must be an int, got ${value.runtimeType}");
  }
}

void _validateIntInRange(int value, int minInclusive, int maxInclusive) {
  _validateInt(value);
  if (value < minInclusive || value > maxInclusive) {
    throw RangeError(
        "value must be in range $minInclusive to $maxInclusive inclusive, got $value");
  }
}

void _validateString(dynamic value) {
  if (value is! String) {
    throw ArgumentError("value must be a String, got ${value.runtimeType}");
  }
}

void _validateStringEnum(String value, Set<String> members) {
  _validateString(value);
  if (!members.contains(value)) {
    throw ArgumentError("value must be one of $members, got '$value'");
  }
}

// --- Base Type Converter Implementations ---

/// Base for integer types, handles parsing and string conversion.
abstract class _BaseIntConverter implements BaseSimpleType<int> {
  @override
  int fromXml(String xmlValue) {
    final value = int.tryParse(xmlValue);
    if (value == null) {
      throw InvalidXmlError("Cannot parse '$xmlValue' as integer");
    }
    // Let subclasses validate the parsed value range if needed
    validate(value);
    return value;
  }

  @override
  String? toXml(int value) {
    validate(value);
    return value.toString();
  }
}

/// Base for string types, handles passthrough conversion.
abstract class _BaseStringConverter implements BaseSimpleType<String> {
  @override
  String fromXml(String xmlValue) {
    return xmlValue;
  }

  @override
  String? toXml(String value) {
    validate(value);
    return value;
  }

  @override
  void validate(String value) {
    _validateString(value);
    // Subclasses can add enum checks etc.
  }
}

/// Base for string enumeration types, adds validation against a set of members.
abstract class _BaseStringEnumConverter extends _BaseStringConverter {
  Set<String> get members;

  @override
  void validate(String value) {
    super.validate(value);
    _validateStringEnum(value, members);
  }
}

// --- XSD Primitive Type Converters ---

class XsdAnyUriConverter extends _BaseStringConverter {
  // No extra validation implemented, as per Python source comment.
}

final xsdAnyUriConverter = XsdAnyUriConverter();

class XsdBooleanConverter implements BaseSimpleType<bool> {
  @override
  bool fromXml(String xmlValue) {
    const trueValues = {"1", "true"};
    const falseValues = {"0", "false"};
    final lowerVal = xmlValue.toLowerCase();

    if (trueValues.contains(lowerVal)) return true;
    if (falseValues.contains(lowerVal)) return false;

    throw InvalidXmlError(
        "value must be one of '1', '0', 'true' or 'false', got '$xmlValue'");
  }

  @override
  String? toXml(bool value) {
    validate(value);
    return value ? "1" : "0";
  }

  @override
  void validate(bool value) {
    // Dart type system handles bool check.
  }
}

final xsdBooleanConverter = XsdBooleanConverter();

class XsdIdConverter extends _BaseStringConverter {
  // Validation omitted as per Python source.
}

final xsdIdConverter = XsdIdConverter();

class XsdIntConverter extends _BaseIntConverter {
  static const int minInclusive = -2147483648;
  static const int maxInclusive = 2147483647;

  @override
  void validate(int value) {
    _validateIntInRange(value, minInclusive, maxInclusive);
  }
}

final xsdIntConverter = XsdIntConverter();

class XsdLongConverter extends _BaseIntConverter {
  // Dart's int handles 64-bit.
  static final int minInclusive = -9223372036854775808;
  static final int maxInclusive = 9223372036854775807;

  @override
  void validate(int value) {
    // No range check needed for standard 64-bit int in Dart
    _validateInt(value);
  }
}

final xsdLongConverter = XsdLongConverter();

class XsdStringConverter extends _BaseStringConverter {}

final xsdStringConverter = XsdStringConverter();

// Note: XsdStringEnumerationConverter is abstract, use specific instances below

class XsdTokenConverter extends _BaseStringConverter {
  // Whitespace collapsing not implemented. Simply treat as string for now.
}

final xsdTokenConverter = XsdTokenConverter();

class XsdUnsignedIntConverter extends _BaseIntConverter {
  static const int minInclusive = 0;
  static const int maxInclusive = 4294967295;

  @override
  void validate(int value) {
    _validateIntInRange(value, minInclusive, maxInclusive);
  }
}

final xsdUnsignedIntConverter = XsdUnsignedIntConverter();

class XsdUnsignedLongConverter extends _BaseIntConverter {
  static const int minInclusive = 0;
  // Dart's int is 64-bit signed. Use max signed value for practical purposes in OOXML.
  static final int maxInclusive = 9223372036854775807;

  @override
  void validate(int value) {
    _validateIntInRange(value, minInclusive, maxInclusive);
  }
}

final xsdUnsignedLongConverter = XsdUnsignedLongConverter();

// --- Specific ST_* Type Converters ---

class ST_BrClearConverter extends _BaseStringEnumConverter {
  @override
  Set<String> get members => const {"none", "left", "right", "all"};
}

final stBrClearConverter = ST_BrClearConverter();

class ST_BrTypeConverter extends _BaseStringEnumConverter {
  @override
  Set<String> get members => const {"page", "column", "textWrapping"};
}

final stBrTypeConverter = ST_BrTypeConverter();

/// Handles coordinate values, often representing EMUs but potentially including universal measures.
class ST_CoordinateConverter implements BaseSimpleType<Length> {
  // Regex to validate if it's potentially a universal measure
  static final _universalMeasurePattern =
      RegExp(r"^-?\d+(\.\d+)?(mm|cm|in|pt|pc|pi)$");
  static final _stUniversalMeasureConverter = ST_UniversalMeasureConverter();
  static final _stCoordinateUnqualifiedConverter =
      ST_CoordinateUnqualifiedConverter();

  @override
  Length fromXml(String xmlValue) {
    if (_universalMeasurePattern.hasMatch(xmlValue)) {
      return _stUniversalMeasureConverter.fromXml(xmlValue);
    }
    // Assume EMU if no unit suffix matches the universal pattern
    final value = int.tryParse(xmlValue);
    if (value == null) {
      throw InvalidXmlError("Cannot parse coordinate '$xmlValue'");
    }
    _stCoordinateUnqualifiedConverter
        .validate(value); // Validate range after parsing
    return Emu(value);
  }

  @override
  String? toXml(Length value) {
    validate(value);
    // Always output as EMU integer string
    return value.emu.toString();
  }

  @override
  void validate(Length value) {
    // Validates the underlying EMU value against the ST_CoordinateUnqualified range
    _stCoordinateUnqualifiedConverter.validate(value.emu);
  }
}

final stCoordinateConverter = ST_CoordinateConverter();

/// Validates the range for coordinate values (assumed EMU).
class ST_CoordinateUnqualifiedConverter extends _BaseIntConverter {
  static const int minInclusive = -27273042329600;
  static const int maxInclusive = 27273042316900;

  @override
  void validate(int value) {
    _validateIntInRange(value, minInclusive, maxInclusive);
  }
}

// No need for instance, only used by ST_CoordinateConverter

class ST_DecimalNumberConverter extends XsdIntConverter {}

final stDecimalNumberConverter = ST_DecimalNumberConverter();

class ST_DrawingElementIdConverter extends XsdUnsignedIntConverter {}

final stDrawingElementIdConverter = ST_DrawingElementIdConverter();

/// Constant container for the 'auto' value used in ST_HexColor.
class ST_HexColorAuto {
  static const String AUTO = "auto";
}

/// Handles hex color strings like "FF0000" or "auto".
/// Converts to/from RGBColor or the special string "auto".
/// Uses `Object` as the type parameter because it can return either RGBColor or String.
class ST_HexColorConverter implements BaseSimpleType<Object> {
  @override
  Object fromXml(String xmlValue) {
    if (xmlValue.toLowerCase() == ST_HexColorAuto.AUTO) {
      return ST_HexColorAuto.AUTO;
    }
    // Use RGBColor's parser which should handle validation
    try {
      return RGBColor.fromString(xmlValue);
    } catch (e) {
      throw InvalidXmlError(
          "Invalid ST_HexColor value '$xmlValue', expected 6-digit hex or 'auto': $e");
    }
  }

  @override
  String? toXml(Object value) {
    validate(value); // Check if it's RGBColor or "auto"
    if (value is RGBColor) {
      return value.toString(); // Assumes RGBColor.toString() produces "RRGGBB"
    } else if (value == ST_HexColorAuto.AUTO) {
      return ST_HexColorAuto.AUTO;
    }
    // Should not happen if validate passes
    throw ArgumentError('Unexpected value type for ST_HexColor: ${value.runtimeType}');
  }

  @override
  void validate(Object value) {
    if (value is RGBColor) {
      return; // RGBColor constructor validates
    }
    if (value is String && value == ST_HexColorAuto.AUTO) {
      return;
    }
    throw ArgumentError(
        "Value must be an RGBColor object or the string '${ST_HexColorAuto.AUTO}', got $value (${value.runtimeType})");
  }
}

final stHexColorConverter = ST_HexColorConverter();


/// Half-point measure, e.g., XML "24" represents 12 points.
class ST_HpsMeasureConverter implements BaseSimpleType<Length> {
  static final _universalMeasurePattern =
      RegExp(r"^(-?\d+(\.\d+)?)(mm|cm|in|pt|pc|pi)$");
  static final _stUniversalMeasureConverter = ST_UniversalMeasureConverter();
  static final _xsdUnsignedLongConverter = XsdUnsignedLongConverter();

  @override
  Length fromXml(String xmlValue) {
    if (_universalMeasurePattern.hasMatch(xmlValue)) {
      return _stUniversalMeasureConverter.fromXml(xmlValue);
    }
    // Assume half-points if no unit suffix
    final halfPoints = int.tryParse(xmlValue);
    if (halfPoints == null) {
      throw InvalidXmlError("Cannot parse half-points value '$xmlValue'");
    }
    _xsdUnsignedLongConverter
        .validate(halfPoints); // Validate range after parsing
    return Pt(halfPoints / 2.0);
  }

  @override
  String? toXml(Length value) {
    validate(value);
    // Convert EMU to half-points
    final halfPoints = (value.pt * 2.0).round();
    return halfPoints.toString();
  }

  @override
  void validate(Length value) {
    // Underlying type is xsdUnsignedLong for the half-point value
    final halfPoints = (value.pt * 2.0).round();
    _xsdUnsignedLongConverter.validate(halfPoints);
  }
}

final stHpsMeasureConverter = ST_HpsMeasureConverter();

class ST_MergeConverter extends _BaseStringEnumConverter {
  static const String continueMerge = "continue";
  static const String restartMerge = "restart";
  @override
  Set<String> get members => const {continueMerge, restartMerge};
}

final stMergeConverter = ST_MergeConverter();

/// Boolean type allowing "on"/"off" in addition to "true"/"false"/"1"/"0".
class ST_OnOffConverter implements BaseSimpleType<bool> {
  @override
  bool fromXml(String xmlValue) {
    const trueValues = {"1", "true", "on"};
    const falseValues = {"0", "false", "off"};
    final lowerVal = xmlValue.toLowerCase();

    if (trueValues.contains(lowerVal)) return true;
    if (falseValues.contains(lowerVal)) return false;

    throw InvalidXmlError(
        "value must be one of '1', '0', 'true', 'false', 'on', or 'off', got '$xmlValue'");
  }

  @override
  String? toXml(bool value) {
    validate(value);
    // Use "1" and "0" for consistency, though "on"/"off" might also be valid output
    return value ? "1" : "0";
  }

  @override
  void validate(bool value) {/* Dart handles boolean type check */}
}

final stOnOffConverter = ST_OnOffConverter();

/// Positive coordinate value, usually representing EMUs.
class ST_PositiveCoordinateConverter implements BaseSimpleType<Length> {
  static const int minInclusive = 0;
  static const int maxInclusive =
      27273042316900; // Max positive value from spec

  @override
  Length fromXml(String xmlValue) {
    final value = int.tryParse(xmlValue);
    if (value == null) {
      throw InvalidXmlError("Cannot parse positive coordinate '$xmlValue'");
    }
    _validateIntInRange(value, minInclusive, maxInclusive);
    return Emu(value);
  }

  @override
  String? toXml(Length value) {
    validate(value);
    return value.emu.toString();
  }

  @override
  void validate(Length value) {
    _validateIntInRange(value.emu, minInclusive, maxInclusive);
  }
}

final stPositiveCoordinateConverter = ST_PositiveCoordinateConverter();

class ST_RelationshipIdConverter extends XsdStringConverter {}

final stRelationshipIdConverter = ST_RelationshipIdConverter();

/// Signed measurement in twentieths of a point (twips).
class ST_SignedTwipsMeasureConverter implements BaseSimpleType<Length> {
  static final _universalMeasurePattern =
      RegExp(r"^(-?\d+(\.\d+)?)(mm|cm|in|pt|pc|pi)$");
  static final _stUniversalMeasureConverter = ST_UniversalMeasureConverter();
  static final _xsdIntConverter = XsdIntConverter();

  @override
  Length fromXml(String xmlValue) {
    if (_universalMeasurePattern.hasMatch(xmlValue)) {
      return _stUniversalMeasureConverter.fromXml(xmlValue);
    }
    // Assume twips if no unit suffix
    // Use double for parsing to handle potential decimals, round to nearest int
    final twipsDouble = double.tryParse(xmlValue);
    if (twipsDouble == null) {
      throw InvalidXmlError("Cannot parse signed twips value '$xmlValue'");
    }
    final twipsValue = twipsDouble.round();
    _xsdIntConverter.validate(twipsValue); // Validate range
    return Twips(twipsValue);
  }

  @override
  String? toXml(Length value) {
    validate(value);
    return value.twips.toString();
  }

  @override
  void validate(Length value) {
    // Underlying type is xsdInt for the twips value
    _xsdIntConverter.validate(value.twips);
  }
}

final stSignedTwipsMeasureConverter = ST_SignedTwipsMeasureConverter();

class ST_StringConverter extends XsdStringConverter {}

final stStringConverter = ST_StringConverter();

class ST_TblLayoutTypeConverter extends _BaseStringEnumConverter {
  @override
  Set<String> get members => const {"fixed", "autofit"};
}

final stTblLayoutTypeConverter = ST_TblLayoutTypeConverter();

class ST_TblWidthConverter extends _BaseStringEnumConverter {
  @override
  Set<String> get members => const {"auto", "dxa", "nil", "pct"};
}

final stTblWidthConverter = ST_TblWidthConverter();

/// Unsigned measurement in twentieths of a point (twips).
class ST_TwipsMeasureConverter implements BaseSimpleType<Length> {
  static final _universalMeasurePattern =
      RegExp(r"^(-?\d+(\.\d+)?)(mm|cm|in|pt|pc|pi)$");
  static final _stUniversalMeasureConverter = ST_UniversalMeasureConverter();
  static final _xsdUnsignedLongConverter = XsdUnsignedLongConverter();


  @override
  Length fromXml(String xmlValue) {
    if (_universalMeasurePattern.hasMatch(xmlValue)) {
      return _stUniversalMeasureConverter.fromXml(xmlValue);
    }
    // Assume twips if no unit suffix
    final twipsValue = int.tryParse(xmlValue);
    if (twipsValue == null) {
       throw InvalidXmlError("Cannot parse unsigned twips value '$xmlValue'");
    }
    // Use unsigned long validation (practical range)
    _xsdUnsignedLongConverter.validate(twipsValue);
    return Twips(twipsValue);
  }

  @override
  String? toXml(Length value) {
    validate(value);
    return value.twips.toString();
  }

  @override
  void validate(Length value) {
    // Underlying type is xsdUnsignedLong for the twips value
    _xsdUnsignedLongConverter.validate(value.twips);
  }
}

final stTwipsMeasureConverter = ST_TwipsMeasureConverter();

/// Handles measurements with explicit units like "1.5in", "12pt".
class ST_UniversalMeasureConverter implements BaseSimpleType<Length> {
  // Regex to capture value and unit (mm, cm, in, pt, pc, pi)
  static final _universalMeasurePattern =
      RegExp(r"^(-?\d+(\.\d+)?)(mm|cm|in|pt|pc|pi)$");
  static const Map<String, int> _unitMultipliers = {
    "mm": 36000,
    "cm": 360000,
    "in": 914400,
    "pt": 12700,
    "pc": 152400, // 1 pica = 12 points
    "pi": 152400, // alternative abbrev for pica
  };

  @override
  Length fromXml(String xmlValue) {
    final match = _universalMeasurePattern.firstMatch(xmlValue);
    if (match == null || match.groupCount < 3) {
      throw InvalidXmlError("Invalid universal measure format: '$xmlValue'");
    }
    final valueStr = match.group(1)!;
    final unit = match.group(3)!;

    final quantity = double.tryParse(valueStr);
    if (quantity == null) {
       throw InvalidXmlError(
          "Cannot parse numeric part '$valueStr' in '$xmlValue'");
    }
    final multiplier = _unitMultipliers[unit];
    if (multiplier == null) {
       // Should not happen if regex matched, but defensive check
      throw InvalidXmlError("Unknown unit '$unit' in '$xmlValue'");
    }
    // Use round() before casting to int for better precision
    return Emu((quantity * multiplier).round());
  }

  @override
  String? toXml(Length value) {
    // Output format is context-dependent (usually twips or EMU), not derived from input.
    // Generally, these values are read, not written directly using this converter.
    // Output should typically use ST_TwipsMeasureConverter or similar.
    throw UnimplementedError(
        "ST_UniversalMeasure.toXml is not typically used for direct output.");
  }

  @override
  void validate(Length value) {
    // No specific validation beyond being a Length object
  }
}

// No top-level instance for UniversalMeasure as it's mainly for parsing input.

class ST_VerticalAlignRunConverter extends _BaseStringEnumConverter {
  static const String baseline = "baseline";
  static const String superscript = "superscript";
  static const String subscript = "subscript";
  @override
  Set<String> get members => const {baseline, superscript, subscript};
}

final stVerticalAlignRunConverter = ST_VerticalAlignRunConverter();




// Conversor para WD_UNDERLINE
class WD_UNDERLINE_Converter implements BaseSimpleType<WD_UNDERLINE> {
  @override
  WD_UNDERLINE fromXml(String xmlValue) {
    // WD_UNDERLINE usa 'xml_value' no Python para mapeamento
    // Precisamos encontrar o membro do enum pelo valor XML
    return WD_UNDERLINE.values.firstWhere(
      (e) => e.xmlValue == xmlValue,
      orElse: () => throw InvalidXmlError(
        "Unknown WD_UNDERLINE value '$xmlValue'",
      ),
    );
  }

  @override
  String? toXml(WD_UNDERLINE value) {
    validate(value);
    // Usa 'xmlValue' definido no enum
    // Retorna null se o xmlValue for null (como para INHERITED)
    return value.xmlValue;
  }

  @override
  void validate(WD_UNDERLINE value) {
    // O tipo enum já garante a validade do membro
  }
}

final wdUnderlineConverter = WD_UNDERLINE_Converter();

// Conversor para WD_COLOR_INDEX
class WD_COLOR_INDEX_Converter implements BaseSimpleType<WD_COLOR_INDEX> {
  @override
  WD_COLOR_INDEX fromXml(String xmlValue) {
    return WD_COLOR_INDEX.values.firstWhere(
      (e) => e.xmlValue == xmlValue,
      orElse: () => throw InvalidXmlError(
        "Unknown WD_COLOR_INDEX value '$xmlValue'",
      ),
    );
  }

  @override
  String? toXml(WD_COLOR_INDEX value) {
    validate(value);
    return value.xmlValue; // Retorna null se xmlValue for null
  }

  @override
  void validate(WD_COLOR_INDEX value) {
    // O tipo enum já garante a validade do membro
  }
}

final wdColorIndexConverter = WD_COLOR_INDEX_Converter();



class MSO_THEME_COLOR_INDEX_Converter
    implements BaseSimpleType<MSO_THEME_COLOR_INDEX> {
  @override
  MSO_THEME_COLOR_INDEX fromXml(String xmlValue) {
    return MSO_THEME_COLOR_INDEX.values.firstWhere(
      (e) => e.xmlValue == xmlValue,
      orElse: () => throw InvalidXmlError(
        "Unknown MSO_THEME_COLOR_INDEX value '$xmlValue'",
      ),
    );
  }

  @override
  String? toXml(MSO_THEME_COLOR_INDEX value) {
    validate(value);
    // Retorna null se xmlValue for null (como para NOT_THEME_COLOR)
    return value.xmlValue;
  }

  @override
  void validate(MSO_THEME_COLOR_INDEX value) {
    // O tipo enum já garante a validade do membro
  }
}

final msoThemeColorIndexConverter = MSO_THEME_COLOR_INDEX_Converter();
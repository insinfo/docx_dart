// docx/enum/shape.dart

// Corresponds to Python file docx/enum/shape.py

/// Corresponds to WdInlineShapeType enumeration.
///
/// Alias: `WD_INLINE_SHAPE`
///
/// http://msdn.microsoft.com/en-us/library/office/ff192587.aspx
enum WD_INLINE_SHAPE_TYPE {
  CHART(12, 'Chart'),
  LINKED_PICTURE(4, 'Linked Picture'),
  PICTURE(3, 'Picture'),
  SMART_ART(15, 'SmartArt'),
  NOT_IMPLEMENTED(
      -6, 'Not Implemented'); // Note: Negative value used as in Python

  /// The integer value corresponding to the MS API enum member.
  final int intValue;

  /// The documentation string for the enum member.
  final String description;

  const WD_INLINE_SHAPE_TYPE(this.intValue, this.description);

  @override
  String toString() => '$name ($intValue)';
}

/// Alias for [WD_INLINE_SHAPE_TYPE].
typedef WD_INLINE_SHAPE = WD_INLINE_SHAPE_TYPE;

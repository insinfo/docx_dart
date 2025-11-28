import 'package:xml/xml.dart';

import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/oxml/shape.dart';

/// Temporary stub for image-related OOXML factory helpers.
///
/// Implementations will be filled in as the drawing/inline picture pipeline is
/// ported. For now this keeps analyzer happy and provides a single integration
/// point for future work.
class OxmlConstructors {
  static XmlElement newPicInline({
    required int shapeId,
    required String rId,
    required String filename,
    required Length cx,
    required Length cy,
  }) {
    final inline = CT_Inline.newPicInline(shapeId, rId, filename, cx, cy);
    return inline.element;
  }
}

import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/shared.dart';

/// Simplified paragraph alignment for InlineParagraphStyle.
enum ParagraphAlignment {
  left,
  center,
  right,
  justify,
}

/// Simplified underline styles for InlineRunStyle.
enum Underline {
  none,
  single,
}

/// Helper to map the simplified enums back to docx_dart's internal enums.
extension ParagraphAlignmentMapper on ParagraphAlignment {
  WD_PARAGRAPH_ALIGNMENT get toDocx {
    switch (this) {
      case ParagraphAlignment.left:
        return WD_PARAGRAPH_ALIGNMENT.LEFT;
      case ParagraphAlignment.center:
        return WD_PARAGRAPH_ALIGNMENT.CENTER;
      case ParagraphAlignment.right:
        return WD_PARAGRAPH_ALIGNMENT.RIGHT;
      case ParagraphAlignment.justify:
        return WD_PARAGRAPH_ALIGNMENT.JUSTIFY;
    }
  }
}

extension UnderlineMapper on Underline {
  WD_UNDERLINE get toDocx {
    switch (this) {
      case Underline.none:
        return WD_UNDERLINE.NONE;
      case Underline.single:
        return WD_UNDERLINE.SINGLE;
    }
  }
}

/// Helper class for easily generating common Word units.
class DocxUnit {
  /// Returns the equivalent of [value] in centimeters.
  static Length cm(double value) => Cm(value);

  /// Returns the equivalent of [value] in points.
  static Length pt(double value) => Pt(value);
}

/// Defines inline formatting properties applied directly to a Paragraph.
class InlineParagraphStyle {
  final ParagraphAlignment? alignment;
  final double? spaceBeforePt;
  final double? spaceAfterPt;
  final double? lineSpacing;
  final double? firstLineIndentCm;

  const InlineParagraphStyle({
    this.alignment,
    this.spaceBeforePt,
    this.spaceAfterPt,
    this.lineSpacing,
    this.firstLineIndentCm,
  });
}

/// Defines inline formatting properties applied directly to a Run.
class InlineRunStyle {
  final bool bold;
  final bool italic;
  final Underline underline;
  final double? fontSizePt;
  final String? fontFamily;

  const InlineRunStyle({
    this.bold = false,
    this.italic = false,
    this.underline = Underline.none,
    this.fontSizePt,
    this.fontFamily,
  });
}

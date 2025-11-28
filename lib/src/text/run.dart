import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/parts/story.dart';
import 'package:docx_dart/src/oxml/shape.dart';
import 'package:docx_dart/src/oxml/text/run.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/shape.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/types.dart';

/// Proxy object wrapping a `<w:r>` element.
class Run extends StoryChild implements ProvidesXmlPart {
  Run(this._r, ProvidesStoryPart parent) : super(parent);

  final CT_R _r;

  @override
  StoryPart get part => super.part;

  /// Add a break element of [breakType] to this run.
  void addBreak([WD_BREAK breakType = WD_BREAK.LINE]) {
    final mapping = _breakTypeMap[breakType] ?? const _BreakMapping();
    _r.addBr(type: mapping.type, clear: mapping.clear);
  }

  /// Add a picture to the end of this run and return the created inline shape.
  InlineShape addPicture(
    dynamic imagePathOrStream, {
    Length? width,
    Length? height,
  }) {
    final inlineElement = part.newPicInline(
      imagePathOrStream,
      width: width,
      height: height,
    );
    final inline = CT_Inline(inlineElement);
    _r.addDrawing(inline);
    return InlineShape(inline);
  }

  /// Add a tab character to this run.
  void addTab() {
    _r.addTab();
  }

  /// Append a text node to the run.
  void addText(String text) {
    _r.addT(text);
  }

  /// Remove all child content from this run, preserving run properties.
  Run clear() {
    _r.clearContent();
    return this;
  }

  /// Whether one or more rendered page-breaks occur in this run.
  bool get containsPageBreak => _r.lastRenderedPageBreaks.isNotEmpty;

  /// String formed by concatenating the text equivalent of each child element.
  String get text => _r.text;

  set text(String value) => _r.text = value;

  /// Character style applied to this run, if any.
  CharacterStyle? get style {
    final resolved = part.getStyle(_r.style, WD_STYLE_TYPE.CHARACTER);
    return resolved is CharacterStyle ? resolved : null;
  }

  set style(dynamic styleOrName) {
    final styleId = part.getStyleId(styleOrName, WD_STYLE_TYPE.CHARACTER);
    _r.style = styleId;
  }
}

class _BreakMapping {
  const _BreakMapping({this.type, this.clear});
  final String? type;
  final String? clear;
}

const Map<WD_BREAK, _BreakMapping> _breakTypeMap = {
  WD_BREAK.LINE: _BreakMapping(),
  WD_BREAK.PAGE: _BreakMapping(type: 'page'),
  WD_BREAK.COLUMN: _BreakMapping(type: 'column'),
  WD_BREAK.LINE_CLEAR_LEFT: _BreakMapping(type: 'textWrapping', clear: 'left'),
  WD_BREAK.LINE_CLEAR_RIGHT:
      _BreakMapping(type: 'textWrapping', clear: 'right'),
  WD_BREAK.LINE_CLEAR_ALL:
      _BreakMapping(type: 'textWrapping', clear: 'all'),
};

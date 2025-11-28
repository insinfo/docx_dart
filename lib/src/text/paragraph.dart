import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/oxml/text/hyperlink.dart';
import 'package:docx_dart/src/oxml/text/pagebreak.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';
import 'package:docx_dart/src/oxml/text/run.dart';
import 'package:docx_dart/src/parts/story.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/text/hyperlink.dart';
import 'package:docx_dart/src/text/parfmt.dart';
import 'package:docx_dart/src/text/run.dart';
import 'package:docx_dart/src/types.dart';

/// Proxy object wrapping a `<w:p>` element.
class Paragraph extends StoryChild implements ProvidesXmlPart {
	Paragraph(this._p, ProvidesStoryPart parent) : super(parent);

	final CT_P _p;

	@override
	StoryPart get part => super.part;

	/// Append a run optionally containing [text] and styled with [style].
	Run addRun([String? text, dynamic style]) {
		final r = _p.add_r();
		final run = Run(r, parent);
		if (text != null) {
			run.text = text;
		}
		if (style != null) {
			run.style = style;
		}
		return run;
	}

	WD_PARAGRAPH_ALIGNMENT? get alignment => _p.alignment;
	set alignment(WD_PARAGRAPH_ALIGNMENT? value) => _p.alignment = value;

	/// Remove all content from this paragraph while preserving formatting.
	Paragraph clear() {
		_p.clearContent();
		return this;
	}

	/// True when the paragraph contains one or more rendered page-breaks.
	bool get containsPageBreak => _p.lastRenderedPageBreaks.isNotEmpty;

	/// Hyperlinks contained in this paragraph.
	List<Hyperlink> get hyperlinks => _p.hyperlink_lst
			.map((hyperlink) => Hyperlink(hyperlink, parent))
			.toList(growable: false);

	/// Insert a paragraph directly before this one, optionally seeded with content.
	Paragraph insertParagraphBefore([String? text, dynamic style]) {
		final paragraph = _insertParagraphBefore();
		if (text != null) {
			paragraph.addRun(text);
		}
		if (style != null) {
			paragraph.style = style;
		}
		return paragraph;
	}

	/// Yields runs and hyperlinks in document order.
	Iterable<dynamic> iterInnerContent() sync* {
		for (final element in _p.innerContentElements) {
			if (element is CT_R) {
				yield Run(element, parent);
			} else if (element is CT_Hyperlink) {
				yield Hyperlink(element, parent);
			}
		}
	}

	/// Access to line spacing, indentation, etc.
	ParagraphFormat get paragraphFormat => ParagraphFormat(_p);

	/// Rendered page-break proxies contained in this paragraph.
	List<RenderedPageBreak> get renderedPageBreaks => _p.lastRenderedPageBreaks
			.map((lrpb) => RenderedPageBreak(lrpb, parent))
			.toList(growable: false);

	/// Runs contained in this paragraph.
	List<Run> get runs =>
			_p.r_lst.map((r) => Run(r, parent)).toList(growable: false);

	/// Paragraph style applied to this paragraph.
	ParagraphStyle? get style {
		final styleId = _p.style;
		final style = part.getStyle(styleId, WD_STYLE_TYPE.PARAGRAPH);
		return style is ParagraphStyle ? style : null;
	}

	set style(dynamic styleOrName) {
		final styleId = part.getStyleId(styleOrName, WD_STYLE_TYPE.PARAGRAPH);
		_p.style = styleId;
	}

	/// Plain-text representation of this paragraph.
	String get text => _p.text;

	set text(String? value) {
		clear();
		if (value != null) {
			addRun(value);
		}
	}

	Paragraph _insertParagraphBefore() {
		final p = _p.add_p_before();
		return Paragraph(p, parent);
	}
}

/// Proxy for a `<w:lastRenderedPageBreak>` element associated with a paragraph.
class RenderedPageBreak extends Parented {
	RenderedPageBreak(this._lrpb, ProvidesStoryPart parent)
			: _storyParent = parent,
				super(parent);

	final CT_LastRenderedPageBreak _lrpb;
	final ProvidesStoryPart _storyParent;

	Paragraph? get precedingParagraphFragment {
		if (_lrpb.precedesAllContent) {
			return null;
		}
		return Paragraph(_lrpb.precedingFragmentP, _storyParent);
	}

	Paragraph? get followingParagraphFragment {
		if (_lrpb.followsAllContent) {
			return null;
		}
		return Paragraph(_lrpb.followingFragmentP, _storyParent);
	}
}
import 'dart:collection';

import 'package:docx_dart/src/enum/shape.dart';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/oxml/shape.dart';
import 'package:docx_dart/src/parts/story.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/oxml/ns.dart';
import 'package:xml/xml.dart';

/// Sequence of inline shapes contained in a document body.
class InlineShapes extends IterableBase<InlineShape> {
  final CT_Body _body;
  final StoryPart _parent;

  InlineShapes(this._body, StoryPart parent) : _parent = parent;

  StoryPart get part => _parent;

  InlineShape operator [](int index) {
    final inlines = _inlineElements;
    if (index < 0 || index >= inlines.length) {
      throw RangeError.index(index, this, 'index');
    }
    return InlineShape(inlines[index]);
  }

  @override
  Iterator<InlineShape> get iterator =>
      _inlineElements.map((inline) => InlineShape(inline)).iterator;

  @override
  int get length => _inlineElements.length;

  List<CT_Inline> get _inlineElements {
    final tagName = CT_Inline.qnTagName;
    final inlines = <CT_Inline>[];
    for (final node in _body.element.descendants.whereType<XmlElement>()) {
      if (node.name.qualified == tagName) {
        inlines.add(CT_Inline(node));
      }
    }
    return inlines;
  }
}

/// Proxy for a `<wp:inline>` element representing an inline shape.
class InlineShape {
  final CT_Inline _inline;

  InlineShape(this._inline);

  Length get height => _inline.extent.cy;
  set height(Length value) {
    _inline.extent.cy = value;
    final pic = _picture;
    if (pic != null) {
      pic.spPr.cy = value;
    }
  }

  WD_INLINE_SHAPE get type {
    final uri = _inline.graphic.graphicData.uri;
    if (uri == nsmap['pic']) {
      final pic = _picture;
      final blip = pic?.blipFill.blip;
      if (blip != null && blip.link != null) {
        return WD_INLINE_SHAPE.LINKED_PICTURE;
      }
      return WD_INLINE_SHAPE.PICTURE;
    }
    if (uri == nsmap['c']) {
      return WD_INLINE_SHAPE.CHART;
    }
    if (uri == nsmap['dgm']) {
      return WD_INLINE_SHAPE.SMART_ART;
    }
    return WD_INLINE_SHAPE.NOT_IMPLEMENTED;
  }

  Length get width => _inline.extent.cx;
  set width(Length value) {
    _inline.extent.cx = value;
    final pic = _picture;
    if (pic != null) {
      pic.spPr.cx = value;
    }
  }

  CT_Picture? get _picture => _inline.graphic.graphicData.pic;
}

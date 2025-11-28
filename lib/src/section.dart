import 'dart:collection';

import 'package:docx_dart/src/blkcntnr.dart';
import 'package:docx_dart/src/enum/section.dart';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/oxml/section.dart';
import 'package:docx_dart/src/oxml/table.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';
import 'package:docx_dart/src/parts/document.dart';
import 'package:docx_dart/src/parts/hdrftr.dart';
import 'package:docx_dart/src/parts/story.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/table.dart';
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:docx_dart/src/types.dart';

/// Document section, providing access to section and page setup settings.
/// Also provides access to headers and footers.
class Section implements ProvidesStoryPart {
  final CT_SectPr _sectPr;
  final DocumentPart _documentPart;

  late final _Footer _primaryFooter =
      _Footer(_sectPr, _documentPart, WD_HEADER_FOOTER.PRIMARY);
  late final _Header _primaryHeader =
      _Header(_sectPr, _documentPart, WD_HEADER_FOOTER.PRIMARY);

  Section(this._sectPr, this._documentPart);

  Length? get bottomMargin => _sectPr.bottomMargin;
  set bottomMargin(Length? value) => _sectPr.bottomMargin = value;

  bool get differentFirstPageHeaderFooter => _sectPr.titlePg_val;
  set differentFirstPageHeaderFooter(bool value) {
    _sectPr.titlePg_val = value;
  }

  _Footer get evenPageFooter =>
      _Footer(_sectPr, _documentPart, WD_HEADER_FOOTER.EVEN_PAGE);
  _Header get evenPageHeader =>
      _Header(_sectPr, _documentPart, WD_HEADER_FOOTER.EVEN_PAGE);
  _Footer get firstPageFooter =>
      _Footer(_sectPr, _documentPart, WD_HEADER_FOOTER.FIRST_PAGE);
  _Header get firstPageHeader =>
      _Header(_sectPr, _documentPart, WD_HEADER_FOOTER.FIRST_PAGE);

  _Footer get footer => _primaryFooter;
  Length? get footerDistance => _sectPr.footer;
  set footerDistance(Length? value) => _sectPr.footer = value;

  Length? get gutter => _sectPr.gutter;
  set gutter(Length? value) => _sectPr.gutter = value;

  _Header get header => _primaryHeader;
  Length? get headerDistance => _sectPr.header;
  set headerDistance(Length? value) => _sectPr.header = value;

  Iterable<dynamic /* Paragraph | Table */ > iterInnerContent() sync* {
    for (final element in _sectPr.iterInnerContent()) {
      if (element is CT_P) {
        yield Paragraph(element, this);
      } else if (element is CT_Tbl) {
        yield Table(element, this);
      }
    }
  }

  Length? get leftMargin => _sectPr.left_margin;
  set leftMargin(Length? value) => _sectPr.left_margin = value;

  WD_ORIENTATION get orientation => _sectPr.orientation;
  set orientation(WD_ORIENTATION? value) => _sectPr.orientation = value;

  Length? get pageHeight => _sectPr.page_height;
  set pageHeight(Length? value) => _sectPr.page_height = value;

  Length? get pageWidth => _sectPr.page_width;
  set pageWidth(Length? value) => _sectPr.page_width = value;

  @override
  StoryPart get part => _documentPart;

  Length? get rightMargin => _sectPr.right_margin;
  set rightMargin(Length? value) => _sectPr.right_margin = value;

  WD_SECTION_START get startType => _sectPr.start_type;
  set startType(WD_SECTION_START? value) => _sectPr.start_type = value;

  Length? get topMargin => _sectPr.top_margin;
  set topMargin(Length? value) => _sectPr.top_margin = value;
}

/// Sequence of [Section] objects corresponding to the sections in the document.
class Sections extends IterableBase<Section> {
  final CT_Document _documentElm;
  final DocumentPart _documentPart;

  Sections(this._documentElm, this._documentPart);

  @override
  Iterator<Section> get iterator => _sectPrList
      .map((sectPr) => Section(sectPr, _documentPart))
      .iterator;

  Section operator [](int index) => Section(_sectPrList[index], _documentPart);

  int get length => _sectPrList.length;

  List<CT_SectPr> get _sectPrList => _documentElm.sectPr_lst;
}

/// Base class for header and footer classes.
class _BaseHeaderFooter extends BlockItemContainer {
  final CT_SectPr _sectPr;
  final DocumentPart _documentPart;
  final WD_HEADER_FOOTER _hdrftrIndex;

  _BaseHeaderFooter(
    CT_SectPr sectPr,
    DocumentPart documentPart,
    WD_HEADER_FOOTER hdrftrIndex,
  )   : _sectPr = sectPr,
        _documentPart = documentPart,
        _hdrftrIndex = hdrftrIndex,
        super.lazy(
          () => throw StateError('Uninitialized header/footer element'),
          documentPart,
        ) {
    setElementProvider(() => _getOrAddDefinition().element);
  }

  bool get isLinkedToPrevious => !_hasDefinition;
  set isLinkedToPrevious(bool value) {
    final targetState = value;
    if (targetState == isLinkedToPrevious) {
      return;
    }
    if (targetState) {
      _dropDefinition();
    } else {
      _addDefinition();
    }
  }

  @override
  StoryPart get part => _getOrAddDefinition();

    StoryPart _addDefinition() =>
      throw UnimplementedError('Must be implemented by subclass');
    StoryPart get _definition =>
      throw UnimplementedError('Must be implemented by subclass');
  void _dropDefinition() =>
      throw UnimplementedError('Must be implemented by subclass');
  bool get _hasDefinition =>
      throw UnimplementedError('Must be implemented by subclass');
  _BaseHeaderFooter? get _priorHeaderfooter => throw UnimplementedError(
      'Must be implemented by subclass');

  StoryPart _getOrAddDefinition() {
    if (_hasDefinition) {
      return _definition;
    }
    final prior = _priorHeaderfooter;
    if (prior != null) {
      return prior._getOrAddDefinition();
    }
    return _addDefinition();
  }
}

class _Footer extends _BaseHeaderFooter {
  _Footer(CT_SectPr sectPr, DocumentPart documentPart,
      WD_HEADER_FOOTER hdrftrIndex)
      : super(sectPr, documentPart, hdrftrIndex);

  @override
  FooterPart _addDefinition() {
    final (footerPart, rId) = _documentPart.addFooterPart();
    _sectPr.addFooterReference(_hdrftrIndex, rId);
    return footerPart;
  }

  @override
  FooterPart get _definition {
    final footerRef = _sectPr.getFooterReference(_hdrftrIndex);
    if (footerRef == null) {
      throw StateError('Footer definition not found for $_hdrftrIndex');
    }
    return _documentPart.footerPart(footerRef.rId);
  }

  @override
  void _dropDefinition() {
    final rId = _sectPr.removeFooterReference(_hdrftrIndex);
    _documentPart.dropRel(rId);
  }

  @override
  bool get _hasDefinition => _sectPr.getFooterReference(_hdrftrIndex) != null;

  @override
  _Footer? get _priorHeaderfooter {
    final preceding = _sectPr.preceding_sectPr;
    if (preceding == null) {
      return null;
    }
    return _Footer(preceding, _documentPart, _hdrftrIndex);
  }
}

class _Header extends _BaseHeaderFooter {
  _Header(CT_SectPr sectPr, DocumentPart documentPart,
      WD_HEADER_FOOTER hdrftrIndex)
      : super(sectPr, documentPart, hdrftrIndex);

  @override
  HeaderPart _addDefinition() {
    final (headerPart, rId) = _documentPart.addHeaderPart();
    _sectPr.addHeaderReference(_hdrftrIndex, rId);
    return headerPart;
  }

  @override
  HeaderPart get _definition {
    final headerRef = _sectPr.getHeaderReference(_hdrftrIndex);
    if (headerRef == null) {
      throw StateError('Header definition not found for $_hdrftrIndex');
    }
    return _documentPart.headerPart(headerRef.rId);
  }

  @override
  void _dropDefinition() {
    final rId = _sectPr.removeHeaderReference(_hdrftrIndex);
    _documentPart.dropHeaderPart(rId);
  }

  @override
  bool get _hasDefinition => _sectPr.getHeaderReference(_hdrftrIndex) != null;

  @override
  _Header? get _priorHeaderfooter {
    final preceding = _sectPr.preceding_sectPr;
    if (preceding == null) {
      return null;
    }
    return _Header(preceding, _documentPart, _hdrftrIndex);
  }
}

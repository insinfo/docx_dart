import 'package:docx_dart/src/blkcntnr.dart';
import 'package:docx_dart/src/enum/section.dart';
import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/parts/document.dart';
import 'package:docx_dart/src/section.dart';
import 'package:docx_dart/src/settings.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/styles/styles.dart';
import 'package:docx_dart/src/table.dart';
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:docx_dart/src/opc/coreprops.dart';
import 'package:docx_dart/src/shape.dart';
import 'package:docx_dart/src/types.dart';

/// WordprocessingML (WML) document.
///
/// Not intended to be constructed directly. Use `loadDocxDocument()` to open or
/// create a document from a `.docx` package.
class Document extends ElementProxy implements ProvidesStoryPart {
  final CT_Document _element;
  final DocumentPart _part;
  Body? __body;

  Document(this._element, this._part) : super(_element);

  /// Return a heading paragraph newly added to the end of the document.
  ///
  /// The heading paragraph will contain [text] and have its paragraph style
  /// determined by [level]. If [level] is 0, the style is set to `Title`. If [level]
  /// is 1 (or omitted), `Heading 1` is used. Otherwise the style is set to `Heading
  /// {level}`. Raises [ArgumentError] if [level] is outside the range 0-9.
  Paragraph addHeading({String text = "", int level = 1}) {
    if (level < 0 || level > 9) {
      throw ArgumentError("level must be in range 0-9, got $level");
    }
    final style = level == 0 ? "Title" : "Heading $level";
    return addParagraph(text: text, style: style);
  }

  /// Return newly [Paragraph] object containing only a page break.
  Paragraph addPageBreak() {
    final paragraph = addParagraph();
    paragraph.addRun().addBreak(WD_BREAK.PAGE);
    return paragraph;
  }

  /// Return paragraph newly added to the end of the document.
  ///
  /// The paragraph is populated with [text] and having paragraph style [style].
  ///
  /// [text] can contain tab (`\t`) characters, which are converted to the
  /// appropriate XML form for a tab. [text] can also include newline (`\n`) or
  /// carriage return (`\r`) characters, each of which is converted to a line
  /// break.
  Paragraph addParagraph({String text = "", dynamic style /* String | ParagraphStyle | None */}) {
    return _body.addParagraph(text: text, style: style);
  }

  /// Return new picture shape added in its own paragraph at end of the document.
  ///
  /// The picture contains the image at [imagePathOrStream], scaled based on
  /// [width] and [height]. If neither width nor height is specified, the picture
  /// appears at its native size. If only one is specified, it is used to compute a
  /// scaling factor that is then applied to the unspecified dimension, preserving the
  /// aspect ratio of the image. The native size of the picture is calculated using
  /// the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi
  /// if no value is specified, as is often the case.
  InlineShape addPicture(
    dynamic imagePathOrStream, { /* String | IO[bytes] */
    Length? width,
    Length? height,
  }) {
    final run = addParagraph().addRun();
    return run.addPicture(imagePathOrStream, width: width, height: height);
  }

  /// Return a [Section] object newly added at the end of the document.
  ///
  /// The optional [startType] argument must be a member of the [WD_SECTION]
  /// enumeration, and defaults to [WD_SECTION.NEW_PAGE] if not provided.
  Section addSection({WD_SECTION startType = WD_SECTION.NEW_PAGE}) {
    final newSectPr = _element.body.addSectionBreak();
    newSectPr.startType = startType;
    return Section(newSectPr, _part);
  }

  /// Add a table having row and column counts of [rows] and [cols] respectively.
  ///
  /// [style] may be a table style object or a table style name. If [style] is `null`,
  /// the table inherits the default table style of the document.
  Table addTable(int rows, int cols, {dynamic style /* String | _TableStyle | None */}) {
    final table = _body.addTable(rows, cols, _blockWidth);
    table.style = style;
    return table;
  }

  /// A [CoreProperties] object providing Dublin Core properties of document.
  CoreProperties get coreProperties => _part.coreProperties;

  /// The [InlineShapes] collection for this document.
  ///
  /// An inline shape is a graphical object, such as a picture, contained in a run of
  /// text and behaving like a character glyph, being flowed like other text in a
  /// paragraph.
  InlineShapes get inlineShapes => _part.inlineShapes;

  /// Generate each `Paragraph` or `Table` in this document in document order.
  Iterable<dynamic /* Paragraph | Table */> iterInnerContent() {
    return _body.iterInnerContent();
  }

  /// The [Paragraph] instances in the document, in document order.
  ///
  /// Note that paragraphs within revision marks such as `<w:ins>` or `<w:del>` do
  /// not appear in this list.
  List<Paragraph> get paragraphs => _body.paragraphs;

  /// The [DocumentPart] object of this document.
  @override
  DocumentPart get part => _part;

  /// Save this document to [pathOrStream].
  ///
  /// [pathOrStream] can be either a path to a filesystem location (a string) or a
  /// file-like object.
  void save(dynamic pathOrStream) {
    _part.save(pathOrStream);
  }

  /// [Sections] object providing access to each section in this document.
  Sections get sections => Sections(_element, _part);

  /// A [Settings] object providing access to the document-level settings.
  Settings get settings => _part.settings;

  /// A [Styles] object providing access to the styles in this document.
  Styles get styles => _part.styles;

  /// All [Table] instances in the document, in document order.
  ///
  /// Note that only tables appearing at the top level of the document appear in this
  /// list; a table nested inside a table cell does not appear. A table within
  /// revision marks such as `<w:ins>` or `<w:del>` will also not appear in the
  /// list.
  List<Table> get tables => _body.tables;

  /// A [Length] object specifying the space between margins in last section.
  Length get _blockWidth {
    final section = sections.isNotEmpty
        ? sections.last
        : Section(_element.body.getOrAddSectPr(), _part);

    final defaultPageWidth = Inches(8.5).emu;
    final defaultMargin = Inches(1).emu;

    final pageWidth = section.pageWidth?.emu ?? defaultPageWidth;
    final left = section.leftMargin?.emu ?? defaultMargin;
    final right = section.rightMargin?.emu ?? defaultMargin;

    final widthEmu = pageWidth - left - right;
    final safeWidth = widthEmu.clamp(0, pageWidth);
    return Emu(safeWidth.toInt());
  }

  /// The [Body] instance containing the content for this document.
  Body get _body {
    __body ??= Body(_element.body, this);
    return __body!;
  }
}

/// Proxy for `<w:body>` element in this document.
///
/// It's primary role is a container for document content.
class Body extends BlockItemContainer {
  final CT_Body _bodyElement;

  Body(this._bodyElement, Document parent) : super(_bodyElement, parent);

  /// Return this [Body] instance after clearing it of all content.
  ///
  /// Section properties for the main document story, if present, are preserved.
  Body clearContent() {
    _bodyElement.clearContent();
    return this;
  }
}

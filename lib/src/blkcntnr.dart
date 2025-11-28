/// Path: lib/src/blkcntnr.dart
/// Based on python-docx: docx/blkcntnr.py
/// Base class for proxy objects that can contain block items (paragraphs, tables).
import 'dart:core';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/oxml/section.dart';
import 'package:docx_dart/src/oxml/table.dart';
import 'package:docx_dart/src/oxml/text/paragraph.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/text/paragraph.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/table.dart';
import 'package:docx_dart/src/types.dart';

// Assuming these CT_* classes are defined in their respective oxml files
// and have the necessary methods like add_p(), _insert_tbl(), p_lst, tbl_lst,
// inner_content_elements.
import 'oxml/document.dart' show CT_Body;
import 'oxml/section.dart' show CT_HdrFtr;
import 'oxml/table.dart' show CT_Tbl, CT_Tc;
import 'oxml/text/paragraph.dart' show CT_P;
import 'oxml/xmlchemy.dart' show BaseOxmlElement; // For casting _element if needed


import 'text/paragraph.dart' show Paragraph;
import 'styles/style.dart' show ParagraphStyle;
import 'table.dart' show Table;

/// Path: lib/src/blkcntnr.dart
/// Based on python-docx: docx/blkcntnr.py
/// Base class for proxy objects that can contain block items (paragraphs, tables).

// Type Alias for underlying OXML element type (remains dynamic for flexibility)
typedef BlockItemElement = dynamic; // CT_Body | CT_HdrFtr | CT_Tc

/// Base class for proxy objects that can contain block items like paragraphs and tables.
class BlockItemContainer extends StoryChild {
  /// Lazily provides the underlying OOXML element (CT_Body, CT_Tc, CT_HdrFtr).
  late BlockItemElement Function() _elementProvider;

  /// Creates a container associated with [_element] and belonging to [parent].
  /// [parent] must provide access back to the story part (e.g., DocumentPart).
  BlockItemContainer(BlockItemElement element, ProvidesStoryPart parent)
      : super(parent) {
    _elementProvider = (() => element);
  }

  /// Creates a container whose backing element is resolved on demand.
  BlockItemContainer.lazy(
    BlockItemElement Function() elementProvider,
    ProvidesStoryPart parent,
  )   : super(parent) {
    _elementProvider = elementProvider;
  }

  /// Allows subclasses to update how the backing element is resolved.
  void setElementProvider(BlockItemElement Function() elementProvider) {
    _elementProvider = elementProvider;
  }

  BlockItemElement get _element => _elementProvider();

  /// Return paragraph newly added to the end of the content in this container.
  /// [text] is added in a single run. [style] may be a style name or
  /// a [ParagraphStyle] instance.
  Paragraph addParagraph({String text = "", dynamic style}) {
    final paragraph = _addParagraph(); // Gets API-level Paragraph wrapper
    if (text.isNotEmpty) {
      paragraph.addRun(text);
    }
    if (style != null) {
      paragraph.style = style; // Paragraph handles resolving style ids
    }
    return paragraph;
  }

  /// Return a new [Table] object appended to the container.
  ///
  /// [width] must be provided for body-like containers because it controls the
  /// initial column widths. Cell containers override this method to provide a
  /// sensible default when no width is supplied.
  Table addTable(int rows, int cols, [Length? width]) {
    final tableWidth = width;
    if (tableWidth == null) {
      throw ArgumentError('width must be supplied when adding a table here');
    }
    final ctTbl = CT_Tbl.newTbl(rows, cols, tableWidth);

    // --- CORRECTION: Use UnsupportedError for type mismatch ---
    if (_element is CT_Body || _element is CT_Tc || _element is CT_HdrFtr) {
       try {
          // Attempt dynamic call (assumes method like insertTbl exists on CT_* class)
          (_element as dynamic).insertTbl(ctTbl);
       } catch (e, s) {
         print("Error inserting table into ${_element.runtimeType}: $e\n$s");
         throw UnsupportedError(
           "Container type ${_element.runtimeType} does not support table insertion via 'insertTbl' or similar method."
         );
       }
    } else {
      throw UnsupportedError( // Use UnsupportedError for type issue
          "Cannot add table to unsupported container type: ${_element.runtimeType}");
    }
    // --- End Correction ---

    return Table(ctTbl, this); // Wrap CT_Tbl in API Table, passing this container as parent
  }

  /// Generate each `Paragraph` or `Table` object within this container, in document order.
  Iterable<dynamic /* Paragraph | Table */ > iterInnerContent() sync* {
     // --- CORRECTION: Use UnsupportedError ---
    if (_element is BaseOxmlElement) {
      // Assumes `innerContentElements` exists and returns List<BaseOxmlElement>
      final List<BaseOxmlElement> contentElements = (_element as dynamic).innerContentElements;
      for (final element in contentElements) {
        if (element is CT_P) {
          yield Paragraph(element, this); // Pass this container as parent
        } else if (element is CT_Tbl) {
          yield Table(element, this); // Pass this container as parent
        }
      }
    } else {
       throw UnsupportedError( // Use UnsupportedError
          "Cannot iterate inner content on unsupported container type: ${_element.runtimeType}");
    }
     // --- End Correction ---
  }

  /// Read-only list containing the paragraphs in this container, in document order.
  List<Paragraph> get paragraphs {
    // --- CORRECTION: Use UnsupportedError ---
     if (_element is BaseOxmlElement) {
        // Assumes `pList` getter exists and returns List<CT_P>
        final List<CT_P> pElements = (_element as dynamic).pList;
        return pElements.map((p) => Paragraph(p, this)).toList(); // Pass this container as parent
     } else {
        throw UnsupportedError( // Use UnsupportedError
          "Cannot get paragraphs from unsupported container type: ${_element.runtimeType}");
     }
     // --- End Correction ---
  }

  /// Read-only list containing the tables in this container, in document order.
  List<Table> get tables {
    // --- CORRECTION: Use UnsupportedError ---
     if (_element is BaseOxmlElement) {
        // Assumes `tblList` getter exists and returns List<CT_Tbl>
        final List<CT_Tbl> tblElements = (_element as dynamic).tblList;
        return tblElements.map((tbl) => Table(tbl, this)).toList(); // Pass this container as parent
     } else {
       throw UnsupportedError( // Use UnsupportedError
          "Cannot get tables from unsupported container type: ${_element.runtimeType}");
     }
     // --- End Correction ---
  }

  /// Internal helper to add a new `<w:p>` element to the container and wrap it.
  Paragraph _addParagraph() {
     // --- CORRECTION: Use UnsupportedError ---
     if (_element is BaseOxmlElement) {
        // Assumes `addP()` method exists and returns CT_P
        final CT_P newP = (_element as dynamic).addP();
        return Paragraph(newP, this); // Wrap CT_P in API Paragraph, pass this container as parent
     } else {
         throw UnsupportedError( // Use UnsupportedError
          "Cannot add paragraph to unsupported container type: ${_element.runtimeType}");
     }
     // --- End Correction ---
  }
}


// --- Placeholder definitions (should be properly defined elsewhere) ---

// Define ProvidesStoryPart as an abstract class (interface)
// This should likely live in a types.dart file or similar.
// abstract class ProvidesStoryPart {
//   // The actual StoryPart type might vary (e.g., DocumentPart, HeaderPart)
//   // Using dynamic for now, but ideally a specific base type like StoryPart
//   dynamic get part;
// }

// // StoryChild definition from shared.dart (basic version)
// class StoryChild {
//   final ProvidesStoryPart _parent;
//   StoryChild(this._parent);

//   // Provides access to the containing part (DocumentPart, HeaderPart, etc.)
//   dynamic get part => _parent.part;
// }
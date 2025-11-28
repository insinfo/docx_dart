import 'package:docx_dart/src/document.dart';
import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/coreprops.dart';
import 'package:docx_dart/src/oxml/document.dart';
import 'package:docx_dart/src/package.dart';
import 'package:docx_dart/src/parts/hdrftr.dart';
import 'package:docx_dart/src/parts/numbering.dart';
import 'package:docx_dart/src/parts/settings.dart';
import 'package:docx_dart/src/parts/story.dart';
import 'package:docx_dart/src/parts/styles.dart';
import 'package:docx_dart/src/settings.dart';
import 'package:docx_dart/src/shape.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/styles/styles.dart';
import 'package:docx_dart/src/types.dart';

/// Main document part of a WordprocessingML (WML) package, aka a .docx file.
///
/// Acts as broker to other parts such as image, core properties, and style parts. It
/// also acts as a convenient delegate when a mid-document object needs a service
/// involving a remote ancestor. The `Parented.part` property inherited by many content
/// objects provides access to this part object for that purpose.
class DocumentPart extends StoryPart implements ProvidesStoryPart {
  SettingsPart? _settingsPart;
  StylesPart? _stylesPart;
  NumberingPart? _numberingPart;
  InlineShapes? _inlineShapes;
  late final CT_Document _documentElement = CT_Document(element.element);

  DocumentPart(super.partname, super.contentType, super.element, super.package);

  /// Return (footer_part, rId) pair for newly-created footer part.
  (FooterPart, String) addFooterPart() {
    final footerPart = FooterPart.newPart(_wordPackage);
    final rId = relateTo(footerPart, RELATIONSHIP_TYPE.FOOTER);
    return (footerPart, rId);
  }

  /// Return (header_part, rId) pair for newly-created header part.
  (HeaderPart, String) addHeaderPart() {
    final headerPart = HeaderPart.newPart(_wordPackage);
    final rId = relateTo(headerPart, RELATIONSHIP_TYPE.HEADER);
    return (headerPart, rId);
  }

  /// A [CoreProperties] object providing read/write access to the core properties
  /// of this document.
  CoreProperties get coreProperties => _wordPackage.coreProperties;

  /// A [Document] object providing access to the content of this document.
  Document get document => Document(_documentElement, this);

  /// Remove related header part identified by [rId].
  void dropHeaderPart(String rId) {
    dropRel(rId);
  }

  /// Return [FooterPart] related by [rId].
  FooterPart footerPart(String rId) {
    return relatedParts[rId] as FooterPart;
  }

  /// Return the style in this document matching [styleId].
  ///
  /// Returns the default style for [styleType] if [styleId] is `null` or does not
  /// match a defined style of [styleType].
  BaseStyle? getStyle(String? styleId, WD_STYLE_TYPE styleType) {
    return styles.getById(styleId, styleType);
  }

  /// Return the style_id (String) of the style of [styleType] matching
  /// [styleOrName].
  ///
  /// Returns `null` if the style resolves to the default style for [styleType] or if
  /// [styleOrName] is itself `null`. Raises if [styleOrName] is a style of the
  /// wrong type or names a style not present in the document.
  String? getStyleId(dynamic styleOrName, WD_STYLE_TYPE styleType) {
    return styles.getStyleId(styleOrName, styleType);
  }

  /// Return [HeaderPart] related by [rId].
  HeaderPart headerPart(String rId) {
    return relatedParts[rId] as HeaderPart;
  }

  /// The [InlineShapes] instance containing the inline shapes in the document.
  InlineShapes get inlineShapes {
    _inlineShapes ??= InlineShapes(_documentElement.body, this);
    return _inlineShapes!;
  }

  /// A [NumberingPart] object providing access to the numbering definitions for
  /// this document.
  ///
  /// Creates an empty numbering part if one is not present.
  NumberingPart get numberingPart {
    if (_numberingPart == null) {
      try {
        _numberingPart = partRelatedBy(RELATIONSHIP_TYPE.NUMBERING) as NumberingPart;
      } catch (e) {
        // KeyError in python, likely StateError or similar here if not found
        final numberingPart = NumberingPart.newPart();
        relateTo(numberingPart, RELATIONSHIP_TYPE.NUMBERING);
        _numberingPart = numberingPart;
      }
    }
    return _numberingPart!;
  }

  /// Save this document to [pathOrStream], which can be either a path to a
  /// filesystem location (a string) or a file-like object.
  void save(dynamic pathOrStream) {
    _wordPackage.save(pathOrStream);
  }

  /// A [Settings] object providing access to the settings in the settings part of
  /// this document.
  Settings get settings => settingsPart.settings;

  /// A [Styles] object providing access to the styles in the styles part of this
  /// document.
  Styles get styles => stylesPart.styles;

  /// A [SettingsPart] object providing access to the document-level settings for
  /// this document.
  ///
  /// Creates a default settings part if one is not present.
  SettingsPart get settingsPart {
    if (_settingsPart == null) {
      try {
        _settingsPart = partRelatedBy(RELATIONSHIP_TYPE.SETTINGS) as SettingsPart;
      } catch (e) {
        final settingsPart = SettingsPart.defaultPart(_wordPackage);
        relateTo(settingsPart, RELATIONSHIP_TYPE.SETTINGS);
        _settingsPart = settingsPart;
      }
    }
    return _settingsPart!;
  }

  /// Instance of [StylesPart] for this document.
  ///
  /// Creates an empty styles part if one is not present.
  StylesPart get stylesPart {
    if (_stylesPart == null) {
      try {
        _stylesPart = partRelatedBy(RELATIONSHIP_TYPE.STYLES) as StylesPart;
      } catch (e) {
        final stylesPart = StylesPart.defaultPart(_wordPackage);
        relateTo(stylesPart, RELATIONSHIP_TYPE.STYLES);
        _stylesPart = stylesPart;
      }
    }
    return _stylesPart!;
  }

  @override
  StoryPart get part => this;

  Package get _wordPackage {
    final currentPackage = package;
    if (currentPackage is! Package) {
      throw StateError('DocumentPart requires a wordprocessing Package instance.');
    }
    return currentPackage;
  }
}

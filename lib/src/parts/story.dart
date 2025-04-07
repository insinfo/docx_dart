// docx/parts/story.dart
// Corresponds to Python file: python-docx/src/docx/parts/story.py

import 'dart:io'; // For File/Stream types if used for imageDescriptor
import 'dart:math' show max; // For finding the max ID


import 'package:xml/xml.dart'; // XML package
import 'package:collection/collection.dart'; // For firstWhereOrNull

// Assuming these files exist and contain the necessary translated classes/enums
import '../opc/part.dart';
import '../opc/constants.dart' show RELATIONSHIP_TYPE;
import 'package:docx_dart/src/shared.dart';// Expects Length, Image
import '../image/image.dart'; // For Image class definition
import '../package.dart'; // Expects Package class
import 'document.dart'; // Expects DocumentPart class
import 'image.dart'; // Expects ImagePart class
import '../styles/style.dart'; // Expects BaseStyle
import 'package:docx_dart/src/enum/style.dart';// Expects WD_STYLE_TYPE
import '../oxml/oxml_constructors.dart'; // Expects helper for creating inline XML
import '../oxml/ns.dart'; // For namespace constants like ns['wp'], ns['r'] etc.

/// Base class for story parts like DocumentPart, HeaderPart, FooterPart.
///
/// A story part is one that can contain block-level content such as paragraphs
/// and tables. It provides common methods for manipulating content and accessing
/// related parts like images or styles via the document part.
class StoryPart extends XmlPart {
  DocumentPart? _documentPart; // Cache for lazy loading

  StoryPart(super.partname, super.contentType, super.element, super.package);

  /// Returns a tuple `(rId, image)` for the image specified by [imageDescriptor].
  ///
  /// [rId] is the key for the relationship between this part and the image part.
  /// [image] is an [Image] instance providing access to image properties.
  /// Creates the image part and relationship if they don't already exist.
  ///
  /// [imageDescriptor] can be a file path (String) or a Stream<List<int>>/Uint8List.
  (String, Image) getOrAddImage(dynamic imageDescriptor /* String | Stream<List<int>> | Uint8List */) {
    final currentPackage = package; // Access package safely
    if (currentPackage == null) {
      throw StateError('Cannot add image part: StoryPart is not associated with a package.');
    }
    // Assumes currentPackage.getOrAddImagePart exists and handles different descriptor types
    final imagePart = currentPackage.getOrAddImagePart(imageDescriptor);
    final rId = relateTo(imagePart, RELATIONSHIP_TYPE.IMAGE);
    return (rId, imagePart.image);
  }

  /// Returns the style in this document matching [styleId].
  ///
  /// Returns the default style for [styleType] if [styleId] is `null` or
  /// does not match a defined style of [styleType].
  BaseStyle getStyle(String? styleId, WD_STYLE_TYPE styleType) {
    return _documentPartLazy.getStyle(styleId, styleType);
  }

  /// Returns the style_id ([String]) of the style of [styleType] corresponding
  /// to [styleOrName].
  ///
  /// Returns `null` if the style resolves to the default style for [styleType]
  /// or if [styleOrName] is itself `null`. Raises ArgumentError if
  /// [styleOrName] is a style of the wrong type or names a style not present
  /// in the document.
  ///
  /// [styleOrName] can be a style name (String) or a [BaseStyle] object.
  String? getStyleId(dynamic styleOrName /* String? | BaseStyle? */, WD_STYLE_TYPE styleType) {
    return _documentPartLazy.getStyleId(styleOrName, styleType);
  }

  /// Returns a newly created `wp:inline` [XmlElement].
  ///
  /// The element contains the image specified by [imageDescriptor] and is scaled
  /// based on the values of [width] and [height].
  ///
  /// [imageDescriptor] can be a file path (String) or a Stream<List<int>>/Uint8List.
  XmlElement newPicInline(
    dynamic imageDescriptor, { /* String | Stream<List<int>> | Uint8List */
    Length? width,
    Length? height,
  }) {
    final (rId, image) = getOrAddImage(imageDescriptor);
    final (cx, cy) = image.scaledDimensions(width: width, height: height);
    final shapeId = nextId;
    final filename = image.filename;

    // Creates the <wp:inline> XML structure using a helper/factory
    return OxmlConstructors.newPicInline(
      shapeId: shapeId,
      rId: rId,
      filename: filename,
      cx: cx,
      cy: cy,
    );
  }

  /// Next available positive integer id value in this story XML document.
  ///
  /// The value is determined by incrementing the maximum existing integer `@id`
  /// value found across relevant elements (`id`, `r:id`, `wp:docPr/@id`,
  /// `pic:cNvPr/@id`). Gaps are not filled. Returns 1 if no numeric IDs exist.
  int get nextId {
    int maxId = 0;

    // Efficiently find potential ID attributes
    final potentialIdAttributes = ['id', 'r:id', 'wp:id', 'pic:id'];

    for (final descendant in element.descendants.whereType<XmlElement>()) {
      for (final attrName in potentialIdAttributes) {
        String? idStr;
        if (attrName.contains(':')) {
          final parts = attrName.split(':');
          final prefix = parts[0];
          final localName = parts[1];
          final nsUri = ns[prefix]; // Assumes ns map from ns.dart
          if (nsUri != null) {
            idStr = descendant.getAttribute(localName, namespace: nsUri);
          }
        } else {
          idStr = descendant.getAttribute(attrName);
        }

        if (idStr != null) {
          final idVal = int.tryParse(idStr);
          if (idVal != null) {
            maxId = max(maxId, idVal);
          }
        }
      }
    }
    return maxId + 1;
  }

  /// The [DocumentPart] instance for the document this story part belongs to.
  ///
  /// Access is lazy-loaded and cached. Throws [StateError] if the part is
  /// not associated with a package or if the main document part is invalid.
  DocumentPart get _documentPartLazy {
    if (_documentPart == null) {
      final currentPackage = package;
      if (currentPackage == null) {
        throw StateError('Cannot access document part: StoryPart is not associated with a package.');
      }
      final mainPart = currentPackage.mainDocumentPart;
      if (mainPart is! DocumentPart) {
         throw StateError('Package main document part is not a DocumentPart or is null.');
      }
      _documentPart = mainPart;
    }
    return _documentPart!;
  }
}
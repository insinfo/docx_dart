/// Path: lib/src/opc/constants.dart
/// Based on python-docx: docx/opc/constants.py
///
/// Constant values related to the Open Packaging Convention (OPC).
/// In particular it includes content types and relationship types.

/// Content type URIs (like MIME-types) that specify a part's format.
class CONTENT_TYPE {
  static const String BMP = "image/bmp";
  static const String DML_CHART =
      "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
  static const String DML_CHARTSHAPES =
      "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
  static const String DML_DIAGRAM_COLORS =
      "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
  static const String DML_DIAGRAM_DATA =
      "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
  static const String DML_DIAGRAM_LAYOUT =
      "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
  static const String DML_DIAGRAM_STYLE =
      "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml";
  static const String GIF = "image/gif";
  static const String JPEG = "image/jpeg";
  static const String MS_PHOTO = "image/vnd.ms-photo";
  static const String OFC_CUSTOM_PROPERTIES =
      "application/vnd.openxmlformats-officedocument.custom-properties+xml";
  static const String OFC_CUSTOM_XML_PROPERTIES =
      "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
  static const String OFC_DRAWING =
      "application/vnd.openxmlformats-officedocument.drawing+xml";
  static const String OFC_EXTENDED_PROPERTIES =
      "application/vnd.openxmlformats-officedocument.extended-properties+xml";
  static const String OFC_OLE_OBJECT =
      "application/vnd.openxmlformats-officedocument.oleObject";
  static const String OFC_PACKAGE =
      "application/vnd.openxmlformats-officedocument.package";
  static const String OFC_THEME =
      "application/vnd.openxmlformats-officedocument.theme+xml";
  static const String OFC_THEME_OVERRIDE =
      "application/vnd.openxmlformats-officedocument.themeOverride+xml";
  static const String OFC_VML_DRAWING =
      "application/vnd.openxmlformats-officedocument.vmlDrawing";
  static const String OPC_CORE_PROPERTIES =
      "application/vnd.openxmlformats-package.core-properties+xml";
  static const String OPC_DIGITAL_SIGNATURE_CERTIFICATE =
      "application/vnd.openxmlformats-package.digital-signature-certificate";
  static const String OPC_DIGITAL_SIGNATURE_ORIGIN =
      "application/vnd.openxmlformats-package.digital-signature-origin";
  static const String OPC_DIGITAL_SIGNATURE_XMLSIGNATURE =
      "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
  static const String OPC_RELATIONSHIPS =
      "application/vnd.openxmlformats-package.relationships+xml";
  static const String PML_COMMENTS =
      "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
  static const String PML_COMMENT_AUTHORS =
      "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
  static const String PML_HANDOUT_MASTER =
      "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
  static const String PML_NOTES_MASTER =
      "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
  static const String PML_NOTES_SLIDE =
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
  static const String PML_PRESENTATION_MAIN =
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
  static const String PML_PRES_PROPS =
      "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
  static const String PML_PRINTER_SETTINGS =
      "application/vnd.openxmlformats-officedocument.presentationml.printerSettings";
  static const String PML_SLIDE =
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
  static const String PML_SLIDESHOW_MAIN =
      "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";
  static const String PML_SLIDE_LAYOUT =
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
  static const String PML_SLIDE_MASTER =
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
  static const String PML_SLIDE_UPDATE_INFO =
      "application/vnd.openxmlformats-officedocument.presentationml.slideUpdateInfo+xml";
  static const String PML_TABLE_STYLES =
      "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
  static const String PML_TAGS =
      "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
  static const String PML_TEMPLATE_MAIN =
      "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
  static const String PML_VIEW_PROPS =
      "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";
  static const String PNG = "image/png";
  static const String SML_CALC_CHAIN =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
  static const String SML_CHARTSHEET =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
  static const String SML_COMMENTS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
  static const String SML_CONNECTIONS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
  static const String SML_CUSTOM_PROPERTY =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty";
  static const String SML_DIALOGSHEET =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml";
  static const String SML_EXTERNAL_LINK =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
  static const String SML_PIVOT_CACHE_DEFINITION =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
  static const String SML_PIVOT_CACHE_RECORDS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
  static const String SML_PIVOT_TABLE =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
  static const String SML_PRINTER_SETTINGS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
  static const String SML_QUERY_TABLE =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
  static const String SML_REVISION_HEADERS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml";
  static const String SML_REVISION_LOG =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml";
  static const String SML_SHARED_STRINGS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
  static const String SML_SHEET =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  static const String SML_SHEET_MAIN =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
  static const String SML_SHEET_METADATA =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
  static const String SML_STYLES =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
  static const String SML_TABLE =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
  static const String SML_TABLE_SINGLE_CELLS =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml";
  static const String SML_TEMPLATE_MAIN =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
  static const String SML_USER_NAMES =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml";
  static const String SML_VOLATILE_DEPENDENCIES =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml";
  static const String SML_WORKSHEET =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
  static const String TIFF = "image/tiff";
  static const String WML_COMMENTS =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
  static const String WML_DOCUMENT =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  static const String WML_DOCUMENT_GLOSSARY =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
  static const String WML_DOCUMENT_MAIN =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
  static const String WML_ENDNOTES =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
  static const String WML_FONT_TABLE =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
  static const String WML_FOOTER =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
  static const String WML_FOOTNOTES =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
  static const String WML_HEADER =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
  static const String WML_NUMBERING =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
  static const String WML_PRINTER_SETTINGS =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings";
  static const String WML_SETTINGS =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
  static const String WML_STYLES =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
  static const String WML_WEB_SETTINGS =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml";
  static const String XML = "application/xml";
  static const String X_EMF = "image/x-emf";
  static const String X_FONTDATA = "application/x-fontdata";
  static const String X_FONT_TTF = "application/x-font-ttf";
  static const String X_WMF = "image/x-wmf";
}

/// Constant values for OPC XML namespaces.
class NAMESPACE {
  static const String DML_WORDPROCESSING_DRAWING =
      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  static const String OFC_RELATIONSHIPS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  static const String OPC_RELATIONSHIPS =
      "http://schemas.openxmlformats.org/package/2006/relationships";
  static const String OPC_CONTENT_TYPES =
      "http://schemas.openxmlformats.org/package/2006/content-types";
  static const String WML_MAIN =
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
}

/// Specifies the target mode for an OPC relationship (Internal or External).
enum RELATIONSHIP_TARGET_MODE {
  external,
  internal;

  /// Parses a string (typically from XML like "External" or "Internal")
  /// into the corresponding enum value. Defaults to `internal` if
  /// the input is null or doesn't match "external" (case-insensitive).
  static RELATIONSHIP_TARGET_MODE fromString(String? mode) {
    return (mode?.toLowerCase() == 'external') ? external : internal;
  }

  /// Returns the string representation with the first letter capitalized,
  /// matching the standard OPC representation (e.g., "Internal", "External").
  @override
  String toString() => name[0].toUpperCase() + name.substring(1);
}

/// Specifies relationship types for OPC relationships.
class RELATIONSHIP_TYPE {
  static const String AUDIO =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";
  static const String A_F_CHUNK =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
  static const String CALC_CHAIN =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
  static const String CERTIFICATE =
      "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";
  static const String CHART =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  static const String CHARTSHEET =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
  static const String CHART_USER_SHAPES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
  static const String COMMENTS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
  static const String COMMENT_AUTHORS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
  static const String CONNECTIONS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
  static const String CONTROL =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
  static const String CORE_PROPERTIES =
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
  static const String CUSTOM_PROPERTIES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
  static const String CUSTOM_PROPERTY =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty";
  static const String CUSTOM_XML =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
  static const String CUSTOM_XML_PROPS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
  static const String DIAGRAM_COLORS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";
  static const String DIAGRAM_DATA =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
  static const String DIAGRAM_LAYOUT =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";
  static const String DIAGRAM_QUICK_STYLE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";
  static const String DIALOGSHEET =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
  static const String DRAWING =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
  static const String ENDNOTES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
  static const String EXTENDED_PROPERTIES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
  static const String EXTERNAL_LINK =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
  static const String FONT =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
  static const String FONT_TABLE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
  static const String FOOTER =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
  static const String FOOTNOTES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
  static const String GLOSSARY_DOCUMENT =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
  static const String HANDOUT_MASTER =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";
  static const String HEADER =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
  static const String HYPERLINK =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
  static const String IMAGE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  static const String NOTES_MASTER =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
  static const String NOTES_SLIDE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
  static const String NUMBERING =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
  static const String OFFICE_DOCUMENT =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
  static const String OLE_OBJECT =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
  static const String ORIGIN =
      "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
  static const String PACKAGE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
  static const String PIVOT_CACHE_DEFINITION =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
  static const String PIVOT_CACHE_RECORDS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/spreadsheetml/pivotCacheRecords";
  static const String PIVOT_TABLE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
  static const String PRES_PROPS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
  static const String PRINTER_SETTINGS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
  static const String QUERY_TABLE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
  static const String REVISION_HEADERS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders";
  static const String REVISION_LOG =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog";
  static const String SETTINGS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
  static const String SHARED_STRINGS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
  static const String SHEET_METADATA =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
  static const String SIGNATURE =
      "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
  static const String SLIDE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
  static const String SLIDE_LAYOUT =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
  static const String SLIDE_MASTER =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
  static const String SLIDE_UPDATE_INFO =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo";
  static const String STYLES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
  static const String TABLE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
  static const String TABLE_SINGLE_CELLS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";
  static const String TABLE_STYLES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
  static const String TAGS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
  static const String THEME =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
  static const String THEME_OVERRIDE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
  static const String THUMBNAIL =
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
  static const String USERNAMES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames";
  static const String VIDEO =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
  static const String VIEW_PROPS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
  static const String VML_DRAWING =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
  static const String VOLATILE_DEPENDENCIES =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";
  static const String WEB_SETTINGS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
  static const String WORKSHEET_SOURCE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheetSource";
  static const String XML_MAPS =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps";
}
/// Path: lib/src/opc/spec.dart
/// Based on python-docx: docx/opc/spec.py
///
/// Provides mappings that embody aspects of the Open XML spec ISO/IEC 29500,
/// specifically the default content types associated with file extensions.

import 'constants.dart'; 

/// Default content type mappings based on file extension.
/// This list corresponds to the `default_content_types` tuple in python-docx.
final List<(String, String)> defaultContentTypes = [
  // Note: The original Python list has multiple entries for "bin".
  // In a simple list lookup, only the first one encountered might be used
  // depending on the lookup logic. If specific types are needed for "bin",
  // the lookup logic elsewhere will need to be more sophisticated than a
  // simple linear search or direct map lookup based *only* on extension.
  // For now, translating directly, keeping WML first as it's most relevant.
  ("bin", CONTENT_TYPE.WML_PRINTER_SETTINGS),
  // ("bin", CONTENT_TYPE.PML_PRINTER_SETTINGS), // Uncomment if needed
  // ("bin", CONTENT_TYPE.SML_PRINTER_SETTINGS), // Uncomment if needed
  ("bmp", CONTENT_TYPE.BMP),
  ("emf", CONTENT_TYPE.X_EMF),
  ("fntdata", CONTENT_TYPE.X_FONTDATA),
  ("gif", CONTENT_TYPE.GIF),
  ("jpe", CONTENT_TYPE.JPEG), // Note: Less common extension for JPEG
  ("jpeg", CONTENT_TYPE.JPEG),
  ("jpg", CONTENT_TYPE.JPEG),
  ("png", CONTENT_TYPE.PNG),
  ("rels", CONTENT_TYPE.OPC_RELATIONSHIPS),
  ("tif", CONTENT_TYPE.TIFF), // Note: Less common extension for TIFF
  ("tiff", CONTENT_TYPE.TIFF),
  ("wdp", CONTENT_TYPE.MS_PHOTO),
  ("wmf", CONTENT_TYPE.X_WMF),
  ("xlsx", CONTENT_TYPE.SML_SHEET), // Added from original python list
  ("xml", CONTENT_TYPE.XML),
];

// You might consider using a Map for faster lookups if needed,
// but be aware of the duplicate "bin" key issue mentioned above.
// Example using a Map (would only store the *last* value for duplicate keys):
// final Map<String, String> defaultContentTypeMap = Map.fromEntries(
//   defaultContentTypes.map((pair) => MapEntry(pair.$1, pair.$2))
// );
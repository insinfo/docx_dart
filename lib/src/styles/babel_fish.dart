// docx/styles/babel_fish.dart
// Dart port of python-docx/docx/styles/__init__.py BabelFish helper.

/// Translates between UI-visible style names (e.g. "Heading 1") and the
/// internal names stored in styles.xml (e.g. "heading 1").
class BabelFish {
  static final List<(String ui, String internal)> _styleAliases = [
    ('Caption', 'caption'),
    ('Footer', 'footer'),
    ('Header', 'header'),
    ('Heading 1', 'heading 1'),
    ('Heading 2', 'heading 2'),
    ('Heading 3', 'heading 3'),
    ('Heading 4', 'heading 4'),
    ('Heading 5', 'heading 5'),
    ('Heading 6', 'heading 6'),
    ('Heading 7', 'heading 7'),
    ('Heading 8', 'heading 8'),
    ('Heading 9', 'heading 9'),
  ];

  static final Map<String, String> _uiToInternal = {
    for (final (ui, internal) in _styleAliases) ui: internal,
  };

  static final Map<String, String> _internalToUi = {
    for (final (ui, internal) in _styleAliases) internal: ui,
  };

  /// Converts a UI style name such as `Heading 1` to the internal representation
  /// stored in styles.xml, e.g. `heading 1`.
  static String ui2internal(String uiStyleName) =>
      _uiToInternal[uiStyleName] ?? uiStyleName;

  /// Converts an internal style name such as `heading 1` back to the UI name
  /// shown in Word, e.g. `Heading 1`.
  static String internal2ui(String internalStyleName) =>
      _internalToUi[internalStyleName] ?? internalStyleName;
}

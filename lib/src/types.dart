// Em lib/src/types.dart
import 'opc/part.dart' show XmlPart;
import 'parts/story.dart' show StoryPart;

abstract class ProvidesXmlPart {
  XmlPart get part;
}

abstract class ProvidesStoryPart {
  StoryPart get part;
}
import 'package:docx_dart/src/oxml/settings.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/types.dart';

class Settings extends ElementProxy {
  final CT_Settings _settings;

  Settings(CT_Settings element, [ProvidesXmlPart? parent])
      : _settings = element,
        super(element, parent);

  /// True if this document has distinct odd and even page headers and footers.
  ///
  /// Read/write.
  bool get oddAndEvenPagesHeaderFooter => _settings.evenAndOddHeadersVal;

  set oddAndEvenPagesHeaderFooter(bool value) {
    _settings.evenAndOddHeadersVal = value;
  }
}

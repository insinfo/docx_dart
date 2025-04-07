// docx/oxml/settings.dart
import 'package:docx_dart/src/oxml/shared.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';

class CT_Settings extends BaseOxmlElement {
  CT_Settings(super.element);

  CT_OnOff? get evenAndOddHeaders;

  bool get evenAndOddHeaders_val;
  set evenAndOddHeaders_val(bool? value);

  CT_OnOff get_or_add_evenAndOddHeaders();
  void _remove_evenAndOddHeaders();
}
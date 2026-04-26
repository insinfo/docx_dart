import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/oxml.dart' show parse_xml;
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/oxml/numbering.dart';
import 'package:docx_dart/src/package.dart';

class NumberingPart extends XmlPart {
  _NumberingDefinitions? _numberingDefinitions;

  NumberingPart(
      super.partname, super.contentType, super.element, super.package);

  static NumberingPart newPart(Package package) {
    final partname = PackUri('/word/numbering.xml');
    const contentType = CONTENT_TYPE.WML_NUMBERING;
    final element = parse_xml(_defaultNumberingXml);
    return NumberingPart(partname, contentType, element, package);
  }

  _NumberingDefinitions get numberingDefinitions {
    _numberingDefinitions ??= _NumberingDefinitions(_ctNumbering);
    return _numberingDefinitions!;
  }

  CT_Numbering get _ctNumbering {
    final base = element;
    if (base is CT_Numbering) {
      return base;
    }
    return CT_Numbering(base.element);
  }
}

const String _defaultNumberingXml = r'''
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
''';

class _NumberingDefinitions {
  final CT_Numbering _numbering;

  _NumberingDefinitions(this._numbering);

  int get length => _numbering.numList.length;
}

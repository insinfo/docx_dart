import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/oxml/numbering.dart';

class NumberingPart extends XmlPart {
  _NumberingDefinitions? _numberingDefinitions;

  NumberingPart(super.partname, super.contentType, super.element, super.package);

  static NumberingPart newPart() {
    throw UnimplementedError();
  }

  _NumberingDefinitions get numberingDefinitions {
    _numberingDefinitions ??=
        _NumberingDefinitions(_ctNumbering);
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

class _NumberingDefinitions {
  final CT_Numbering _numbering;

  _NumberingDefinitions(this._numbering);

  int get length => _numbering.numList.length;
}

// docx/opc/parts/coreprops.py
import 'dart:typed_data';
import 'dart:io'; // Para IO
import 'package:docx_dart/src/opc/constants.dart';
import 'package:docx_dart/src/opc/coreprops.dart';
import 'package:docx_dart/src/opc/package.dart';
import 'package:docx_dart/src/opc/packuri.dart';
import 'package:docx_dart/src/opc/part.dart';
import 'package:docx_dart/src/oxml/coreprops.dart';

class CorePropertiesPart extends XmlPart {

 CorePropertiesPart(PackUri partname, String contentType, CT_CoreProperties element, OpcPackage package)
    : super(partname, contentType, element, package);

 /// Cria uma parte de propriedades do núcleo padrão.
 static CorePropertiesPart defaultPart(OpcPackage package) {
    final partname = PackUri("/docProps/core.xml");
    final contentType = CONTENT_TYPE.OPC_CORE_PROPERTIES;
    final corePropertiesElement = CT_CoreProperties.newCoreProperties();

    final part = CorePropertiesPart(partname, contentType, corePropertiesElement, package);

    // Definir valores padrão
    part.coreProperties.title = "Word Document";
    part.coreProperties.lastModifiedBy = "docx_dart"; // Ou similar
    part.coreProperties.revision = 1;
    part.coreProperties.modified = DateTime.now().toUtc();

    return part;
 }

 CoreProperties get coreProperties => CoreProperties(_element as CT_CoreProperties);

 // Método load herdado de XmlPart
}
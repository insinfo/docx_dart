// docx/opc/coreprops.py
import 'package:docx_dart/src/oxml/coreprops.dart';

/// Proxy exposing Dublin Core properties stored in `/docProps/core.xml`.
class CoreProperties {
  final CT_CoreProperties _element;

  CoreProperties(this._element);

  String get author => _element.authorText;
  set author(String value) => _element.authorText = value;

  String get category => _element.categoryText;
  set category(String value) => _element.categoryText = value;

  String get comments => _element.commentsText;
  set comments(String value) => _element.commentsText = value;

  String get contentStatus => _element.contentStatusText;
  set contentStatus(String value) => _element.contentStatusText = value;

  DateTime? get created => _element.createdDatetime;
  set created(DateTime? value) => _element.createdDatetime = value;

  String get identifier => _element.identifierText;
  set identifier(String value) => _element.identifierText = value;

  String get keywords => _element.keywordsText;
  set keywords(String value) => _element.keywordsText = value;

  String get language => _element.languageText;
  set language(String value) => _element.languageText = value;

  String get lastModifiedBy => _element.lastModifiedByText;
  set lastModifiedBy(String value) => _element.lastModifiedByText = value;

  DateTime? get lastPrinted => _element.lastPrintedDatetime;
  set lastPrinted(DateTime? value) => _element.lastPrintedDatetime = value;

  DateTime? get modified => _element.modifiedDatetime;
  set modified(DateTime? value) => _element.modifiedDatetime = value;

  int get revision => _element.revisionNumber;
  set revision(int value) => _element.revisionNumber = value;

  String get subject => _element.subjectText;
  set subject(String value) => _element.subjectText = value;

  String get title => _element.titleText;
  set title(String value) => _element.titleText = value;

  String get version => _element.versionText;
  set version(String value) => _element.versionText = value;
}
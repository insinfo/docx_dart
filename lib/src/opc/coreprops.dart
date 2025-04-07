// docx/opc/coreprops.py
import 'package:docx_dart/src/oxml/coreprops.dart'; // Para CT_CoreProperties

class CoreProperties {
  final CT_CoreProperties _element;

  CoreProperties(this._element);

  String get author;
  set author(String value);

  String get category;
  set category(String value);

  String get comments;
  set comments(String value);

  String get contentStatus;
  set contentStatus(String value);

  DateTime? get created;
  set created(DateTime? value);

  String get identifier;
  set identifier(String value);

  String get keywords;
  set keywords(String value);

  String get language;
  set language(String value);

  String get lastModifiedBy;
  set lastModifiedBy(String value);

  DateTime? get lastPrinted;
  set lastPrinted(DateTime? value);

  DateTime? get modified;
  set modified(DateTime? value);

  int get revision;
  set revision(int value);

  String get subject;
  set subject(String value);

  String get title;
  set title(String value);

  String get version;
  set version(String value);
}
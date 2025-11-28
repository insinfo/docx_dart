// docx/styles/style.dart
// Dart port of python-docx/docx/styles/style.py

import 'package:collection/collection.dart';
import 'package:xml/xml.dart';

import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/oxml/ns.dart';
import 'package:docx_dart/src/oxml/xmlchemy.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/styles/babel_fish.dart';
import 'package:docx_dart/src/types.dart';

class StyleFactory {
	static BaseStyle create(BaseOxmlElement element, ProvidesXmlPart? parent) {
		switch (_styleType(element)) {
			case WD_STYLE_TYPE.PARAGRAPH:
				return ParagraphStyle(element, parent);
			case WD_STYLE_TYPE.CHARACTER:
				return CharacterStyle(element, parent);
			case WD_STYLE_TYPE.TABLE:
				return TableStyle(element, parent);
			case WD_STYLE_TYPE.LIST:
				return _NumberingStyle(element, parent);
		}
	}

	static WD_STYLE_TYPE _styleType(BaseOxmlElement element) {
		final raw = element.element.getAttribute('type', namespace: nsmap['w']);
		if (raw == null) {
			return WD_STYLE_TYPE.PARAGRAPH;
		}
		return WD_STYLE_TYPE.fromXml(raw);
	}

	static String styleIdFromStyleName(String name) {
		final sanitized = name
				.split('')
				.where((ch) => RegExp(r'[A-Za-z0-9_]').hasMatch(ch))
				.join();
		if (sanitized.isEmpty) {
			return 'Style${DateTime.now().microsecondsSinceEpoch}';
		}
		if (RegExp(r'^[0-9]').hasMatch(sanitized)) {
			return '_$sanitized';
		}
		return sanitized;
	}
}

class ParagraphStyle extends BaseStyle {
	ParagraphStyle(BaseOxmlElement element, [ProvidesXmlPart? parent])
			: super(element, parent);
}

class CharacterStyle extends BaseStyle {
	CharacterStyle(BaseOxmlElement element, [ProvidesXmlPart? parent])
			: super(element, parent);
}

class TableStyle extends ParagraphStyle {
	TableStyle(BaseOxmlElement element, [ProvidesXmlPart? parent])
			: super(element, parent);
}

class _NumberingStyle extends BaseStyle {
	_NumberingStyle(BaseOxmlElement element, [ProvidesXmlPart? parent])
			: super(element, parent);
}

class BaseStyle extends ElementProxy {
	BaseStyle(BaseOxmlElement element, [ProvidesXmlPart? parent])
			: super(element, parent);

	XmlElement get _xml => element.element;

	bool get builtin =>
			_xml.getAttribute('customStyle', namespace: nsmap['w']) != '1';

	bool get hidden => _getBoolAttr('semiHidden');
	set hidden(bool value) => _setBoolAttr('semiHidden', value);

	bool get locked => _getBoolAttr('locked');
	set locked(bool value) => _setBoolAttr('locked', value);

	String? get name {
		final nameEl =
				_xml.findElements('name', namespace: nsmap['w']).firstOrNull;
		final value = nameEl?.getAttribute('val', namespace: nsmap['w']);
		return value == null ? null : BabelFish.internal2ui(value);
	}

	set name(String? value) {
		if (value == null) {
			_removeChild('name');
			return;
		}
		final internal = BabelFish.ui2internal(value);
		final existing =
				_xml.findElements('name', namespace: nsmap['w']).firstOrNull;
		if (existing != null) {
			existing.setAttribute('val', internal, namespace: nsmap['w']);
		} else {
			final child = XmlElement(XmlName('name', 'w'));
			child.setAttribute('val', internal, namespace: nsmap['w']);
			_xml.children.add(child);
		}
	}

	int? get priority =>
			_getIntAttr('uiPriority');
	set priority(int? value) => _setIntAttr('uiPriority', value);

	bool get quickStyle => _getBoolAttr('qFormat');
	set quickStyle(bool value) => _setBoolAttr('qFormat', value);

	String? get styleId =>
			_xml.getAttribute('styleId', namespace: nsmap['w']);
	set styleId(String? value) {
		if (value == null) {
			_xml.removeAttribute('styleId', namespace: nsmap['w']);
		} else {
			_xml.setAttribute('styleId', value, namespace: nsmap['w']);
		}
	}

	WD_STYLE_TYPE get type => StyleFactory._styleType(element);

	bool get unhideWhenUsed => _getBoolAttr('unhideWhenUsed');
	set unhideWhenUsed(bool value) => _setBoolAttr('unhideWhenUsed', value);

	void delete() {
		_xml.parent?.children.remove(_xml);
	}

	bool _getBoolAttr(String name) {
		final raw = _xml.getAttribute(name, namespace: nsmap['w']);
		if (raw == null) {
			return false;
		}
		final normalized = raw.toLowerCase();
		return normalized == '1' || normalized == 'true';
	}

	void _setBoolAttr(String name, bool value) {
		if (value) {
			_xml.setAttribute(name, '1', namespace: nsmap['w']);
		} else {
			_xml.removeAttribute(name, namespace: nsmap['w']);
		}
	}

	int? _getIntAttr(String name) {
		final raw = _xml.getAttribute(name, namespace: nsmap['w']);
		return raw == null ? null : int.tryParse(raw);
	}

	void _setIntAttr(String name, int? value) {
		if (value == null) {
			_xml.removeAttribute(name, namespace: nsmap['w']);
		} else {
			_xml.setAttribute(name, value.toString(), namespace: nsmap['w']);
		}
	}

	void _removeChild(String localName) {
		final child = _xml.findElements(localName, namespace: nsmap['w']).firstOrNull;
		if (child != null) {
			_xml.children.remove(child);
		}
	}
}
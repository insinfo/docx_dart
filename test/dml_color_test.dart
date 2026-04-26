import 'package:docx_dart/src/dml/color.dart';
import 'package:docx_dart/src/enum/dml.dart';
import 'package:docx_dart/src/oxml/text/run.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('ColorFormat', () {
    for (final fixture in [
      (_runXml(), null),
      (_runXml(rPr: true), null),
      (_runXml(colorAttrs: {'w:val': 'auto'}), null),
      (
        _runXml(colorAttrs: {'w:val': '4224FF'}),
        const RGBColor(0x42, 0x24, 0xFF)
      ),
      (_runXml(colorAttrs: {'w:val': 'auto', 'w:themeColor': 'accent1'}), null),
      (
        _runXml(colorAttrs: {'w:val': 'F00BA9', 'w:themeColor': 'accent1'}),
        const RGBColor(0xF0, 0x0B, 0xA9)
      ),
    ]) {
      final (runXml, expectedValue) = fixture;

      test('rgb reads ${expectedValue ?? 'null'} from ${_caseName(runXml)}',
          () {
        expect(_colorFormat(runXml).rgb, expectedValue);
      });
    }

    for (final fixture in [
      (_runXml(), const RGBColor(10, 20, 30), '0A141E', null),
      (_runXml(rPr: true), const RGBColor(1, 2, 3), '010203', null),
      (
        _runXml(colorAttrs: {'w:val': '123abc'}),
        const RGBColor(42, 24, 99),
        '2A1863',
        null
      ),
      (
        _runXml(colorAttrs: {'w:val': 'auto'}),
        const RGBColor(16, 17, 18),
        '101112',
        null
      ),
      (
        _runXml(colorAttrs: {'w:val': '234bcd', 'w:themeColor': 'dark1'}),
        const RGBColor(24, 42, 99),
        '182A63',
        null
      ),
      (
        _runXml(colorAttrs: {'w:val': '234bcd', 'w:themeColor': 'dark1'}),
        null,
        null,
        null
      ),
      (_runXml(), null, null, null),
    ]) {
      final (runXml, newValue, expectedVal, expectedTheme) = fixture;

      test('rgb writes ${newValue ?? 'null'} for ${_caseName(runXml)}', () {
        final colorFormat = _colorFormat(runXml);

        colorFormat.rgb = newValue;

        expect(_colorVal(colorFormat), expectedVal);
        expect(_themeColor(colorFormat), expectedTheme);
      });
    }

    for (final fixture in [
      (_runXml(), null),
      (_runXml(rPr: true), null),
      (_runXml(colorAttrs: {'w:val': 'auto'}), null),
      (_runXml(colorAttrs: {'w:val': '4224FF'}), null),
      (
        _runXml(colorAttrs: {'w:themeColor': 'accent1'}),
        MSO_THEME_COLOR.ACCENT_1
      ),
      (
        _runXml(colorAttrs: {'w:val': 'F00BA9', 'w:themeColor': 'dark1'}),
        MSO_THEME_COLOR.DARK_1
      ),
    ]) {
      final (runXml, expectedValue) = fixture;

      test(
          'themeColor reads ${expectedValue ?? 'null'} from ${_caseName(runXml)}',
          () {
        expect(_colorFormat(runXml).themeColor, expectedValue);
      });
    }

    for (final fixture in [
      (_runXml(), MSO_THEME_COLOR.ACCENT_1, '000000', 'accent1'),
      (_runXml(rPr: true), MSO_THEME_COLOR.ACCENT_2, '000000', 'accent2'),
      (
        _runXml(colorAttrs: {'w:val': '101112'}),
        MSO_THEME_COLOR.ACCENT_3,
        '101112',
        'accent3'
      ),
      (
        _runXml(colorAttrs: {'w:val': '234bcd', 'w:themeColor': 'dark1'}),
        MSO_THEME_COLOR.LIGHT_2,
        '234bcd',
        'light2'
      ),
      (
        _runXml(colorAttrs: {'w:val': '234bcd', 'w:themeColor': 'dark1'}),
        null,
        null,
        null
      ),
      (_runXml(), null, null, null),
    ]) {
      final (runXml, newValue, expectedVal, expectedTheme) = fixture;

      test('themeColor writes ${newValue ?? 'null'} for ${_caseName(runXml)}',
          () {
        final colorFormat = _colorFormat(runXml);

        colorFormat.themeColor = newValue;

        expect(_colorVal(colorFormat), expectedVal);
        expect(_themeColor(colorFormat), expectedTheme);
      });
    }

    for (final fixture in [
      (_runXml(), null),
      (_runXml(rPr: true), null),
      (_runXml(colorAttrs: {'w:val': 'auto'}), MSO_COLOR_TYPE.AUTO),
      (_runXml(colorAttrs: {'w:val': '4224FF'}), MSO_COLOR_TYPE.RGB),
      (_runXml(colorAttrs: {'w:themeColor': 'dark1'}), MSO_COLOR_TYPE.THEME),
      (
        _runXml(colorAttrs: {'w:val': 'F00BA9', 'w:themeColor': 'accent1'}),
        MSO_COLOR_TYPE.THEME
      ),
    ]) {
      final (runXml, expectedValue) = fixture;

      test('type reads ${expectedValue ?? 'null'} from ${_caseName(runXml)}',
          () {
        expect(_colorFormat(runXml).type, expectedValue);
      });
    }
  });
}

const _wNamespace =
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

ColorFormat _colorFormat(String runXml) => ColorFormat(_run(runXml));

String? _colorVal(ColorFormat colorFormat) =>
    _colorElement(colorFormat)?.getAttribute('val', namespace: _wNamespace);

String? _themeColor(ColorFormat colorFormat) => _colorElement(colorFormat)
    ?.getAttribute('themeColor', namespace: _wNamespace);

XmlElement? _colorElement(ColorFormat colorFormat) =>
    colorFormat.element.element.descendants
        .whereType<XmlElement>()
        .where((element) => element.name.local == 'color')
        .firstOrNull;

CT_R _run(String runXml) => CT_R(XmlDocument.parse(runXml).rootElement);

String _runXml({bool rPr = false, Map<String, String>? colorAttrs}) {
  final attrs = colorAttrs?.entries
          .map((entry) => '${entry.key}="${entry.value}"')
          .join(' ') ??
      '';
  final colorXml = colorAttrs == null ? '' : '<w:color $attrs/>';
  final rPrXml = rPr || colorAttrs != null ? '<w:rPr>$colorXml</w:rPr>' : '';
  return '<w:r xmlns:w="$_wNamespace">$rPrXml</w:r>';
}

String _caseName(String runXml) => runXml
    .replaceAll(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w')
    .replaceAll(RegExp(r'\s+'), ' ');

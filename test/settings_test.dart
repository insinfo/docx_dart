import 'package:docx_dart/src/oxml/settings.dart';
import 'package:docx_dart/src/settings.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('Settings', () {
    for (final fixture in [
      (_settingsXml(), false),
      (_settingsXml(evenAndOddHeaders: true), true),
      (_settingsXml(evenAndOddHeaders: true, val: '0'), false),
      (_settingsXml(evenAndOddHeaders: true, val: '1'), true),
      (_settingsXml(evenAndOddHeaders: true, val: 'true'), true),
    ]) {
      final (settingsXml, expectedValue) = fixture;

      test('oddAndEvenPagesHeaderFooter reads $expectedValue', () {
        final settings = Settings(_settings(settingsXml));

        expect(settings.oddAndEvenPagesHeaderFooter, expectedValue);
      });
    }

    for (final fixture in [
      (_settingsXml(), true, true),
      (_settingsXml(evenAndOddHeaders: true), false, false),
      (_settingsXml(evenAndOddHeaders: true, val: '1'), true, true),
      (_settingsXml(evenAndOddHeaders: true, val: 'off'), false, false),
    ]) {
      final (settingsXml, value, expectedPresent) = fixture;

      test('oddAndEvenPagesHeaderFooter writes $value', () {
        final settingsElement = _settings(settingsXml);
        final settings = Settings(settingsElement);

        settings.oddAndEvenPagesHeaderFooter = value;

        expect(_hasEvenAndOddHeaders(settingsElement), expectedPresent);
        if (expectedPresent) {
          expect(_evenAndOddHeadersVal(settingsElement), isNull);
        }
      });
    }
  });
}

const _wNamespace =
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

CT_Settings _settings(String settingsXml) =>
    CT_Settings(XmlDocument.parse(settingsXml).rootElement);

bool _hasEvenAndOddHeaders(CT_Settings settings) => settings.element
    .findElements('evenAndOddHeaders', namespace: _wNamespace)
    .isNotEmpty;

String? _evenAndOddHeadersVal(CT_Settings settings) => settings.element
    .findElements('evenAndOddHeaders', namespace: _wNamespace)
    .firstOrNull
    ?.getAttribute('val', namespace: _wNamespace);

String _settingsXml({bool evenAndOddHeaders = false, String? val}) {
  final child = evenAndOddHeaders
      ? '<w:evenAndOddHeaders${val == null ? '' : ' w:val="$val"'}/>'
      : '';
  return '<w:settings xmlns:w="$_wNamespace">$child</w:settings>';
}

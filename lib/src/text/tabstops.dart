import 'dart:collection';

import 'package:docx_dart/src/enum/text.dart';
import 'package:docx_dart/src/oxml/text/parfmt.dart';
import 'package:docx_dart/src/shared.dart';

import 'package:xml/xml.dart';

/// Collection proxy exposing paragraph or style tab stops.
class TabStops extends IterableBase<TabStop> {
  TabStops(this._pPr);

  final CT_PPr _pPr;

  CT_TabStops? get _tabs => _pPr.tabs;

  @override
  Iterator<TabStop> get iterator {
    final tabs = _tabs;
    final elements = tabs?.tab_lst ?? <CT_TabStop>[];
    return elements.map((tab) => TabStop(tab)).iterator;
  }

  @override
  int get length => _tabs?.tab_lst.length ?? 0;

  TabStop operator [](int index) {
    final tabs = _tabs;
    if (tabs == null) {
      throw RangeError.index(index, this, 'index');
    }
    final tabList = tabs.tab_lst;
    if (index < 0 || index >= tabList.length) {
      throw RangeError.index(index, this, 'index');
    }
    return TabStop(tabList[index]);
  }

  /// Remove the tab stop at [index].
  void removeAt(int index) {
    final tabs = _tabs;
    if (tabs == null) {
      throw RangeError.index(index, this, 'index');
    }
    final tabList = tabs.tab_lst;
    if (index < 0 || index >= tabList.length) {
      throw RangeError.index(index, this, 'index');
    }
    final tabElement = tabList[index];
    tabs.element.children.remove(tabElement.element);
    if (tabs.tab_lst.isEmpty) {
      _pPr.removeTabs();
    }
  }

  /// Add a new tab stop ordered by position.
  TabStop addTabStop(
    Length position, {
    WD_TAB_ALIGNMENT alignment = WD_TAB_ALIGNMENT.LEFT,
    WD_TAB_LEADER leader = WD_TAB_LEADER.SPACES,
  }) {
    final tabs = _pPr.getOrAddTabs();
    final tab = tabs.insertTabInOrder(position, alignment, leader);
    return TabStop(tab);
  }

  /// Remove all custom tab stops.
  void clearAll() {
    final tabs = _tabs;
    if (tabs == null) {
      return;
    }
    _pPr.removeTabs();
  }
}

/// Proxy for a single `<w:tab>` definition.
class TabStop extends ElementProxy {
  TabStop(this._tab) : super(_tab);

  CT_TabStop _tab;

  WD_TAB_ALIGNMENT get alignment => _tab.val;
  set alignment(WD_TAB_ALIGNMENT value) => _tab.val = value;

  WD_TAB_LEADER get leader => _tab.leader;
  set leader(WD_TAB_LEADER value) => _tab.leader = value;

  Length get position => _tab.pos;
  set position(Length value) {
    final tabsElement = _tab.element.parent;
    if (tabsElement == null || tabsElement is! XmlElement) {
      _tab.pos = value;
      return;
    }
    final tabs = CT_TabStops(tabsElement);
    final replacement = tabs.insertTabInOrder(value, _tab.val, _tab.leader);
    tabs.element.children.remove(_tab.element);
    _tab = replacement;
  }
}

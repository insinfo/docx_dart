// docx/table.dart
// Ported from python-docx/docx/table.py

import 'dart:collection';

import 'package:docx_dart/src/blkcntnr.dart';
import 'package:docx_dart/src/enum/style.dart';
import 'package:docx_dart/src/enum/table.dart';
import 'package:docx_dart/src/oxml/table.dart';
import 'package:docx_dart/src/shared.dart';
import 'package:docx_dart/src/styles/style.dart';
import 'package:docx_dart/src/types.dart';

class Table extends StoryChild {
	Table(this._tbl, ProvidesStoryPart parent) : super(parent);

	final CT_Tbl _tbl;
	late final _Columns _columns = _Columns(_tbl, this);
	late final _Rows _rows = _Rows(_tbl, this);

	_Column addColumn(Length width) {
		final tblGrid = _tbl.tblGrid;
		final gridCol = tblGrid.addGridCol();
		gridCol.w = width;
		for (final row in _tbl.trList) {
			final tc = row.addTc();
			tc.width = width;
		}
		return _Column(gridCol, this);
	}

	_Row addRow() {
		final tr = _tbl.addTr();
		for (final gridCol in _tbl.tblGrid.gridColList) {
			final tc = tr.addTc();
			final colWidth = gridCol.w;
			if (colWidth != null) {
				tc.width = colWidth;
			}
		}
		return _Row(tr, this);
	}

	WD_TABLE_ALIGNMENT? get alignment => _tblPr.alignment;
	set alignment(WD_TABLE_ALIGNMENT? value) => _tblPr.alignment = value;

	bool get autofit => _tblPr.autofit;
	set autofit(bool value) => _tblPr.autofit = value;

	_Cell cell(int rowIdx, int colIdx) {
		_ensureValidRow(rowIdx);
		_ensureValidColumn(colIdx);
		final columnCount = _columnCount;
		if (columnCount == 0) {
			throw StateError('table has no columns');
		}
		final index = colIdx + (rowIdx * columnCount);
		final cells = _cells;
		if (index < 0 || index >= cells.length) {
			throw RangeError.index(colIdx, cells, 'colIdx');
		}
		return cells[index];
	}

	List<_Cell> columnCells(int columnIdx) {
		_ensureValidColumn(columnIdx);
		final columnCount = _columnCount;
		final cells = _cells;
		final result = <_Cell>[];
		for (var idx = columnIdx; idx < cells.length; idx += columnCount) {
			result.add(cells[idx]);
		}
		return List<_Cell>.unmodifiable(result);
	}

	_Columns get columns => _columns;

	List<_Cell> rowCells(int rowIdx) {
		_ensureValidRow(rowIdx);
		final columnCount = _columnCount;
		final start = rowIdx * columnCount;
		final end = start + columnCount;
		final cells = _cells;
		return List<_Cell>.unmodifiable(cells.sublist(start, end));
	}

	_Rows get rows => _rows;

	TableStyle? get style {
		final styleId = _tbl.tblStyleVal;
		final resolved = part.getStyle(styleId, WD_STYLE_TYPE.TABLE);
		return resolved is TableStyle ? resolved : null;
	}

	set style(Object? styleOrName) {
		final styleId = part.getStyleId(styleOrName, WD_STYLE_TYPE.TABLE);
		_tbl.tblStyleVal = styleId;
	}

	Table get table => this;

	WD_TABLE_DIRECTION? get tableDirection => _tbl.bidiVisualVal;
	set tableDirection(WD_TABLE_DIRECTION? value) => _tbl.bidiVisualVal = value;

	List<_Cell> get _cells {
		final columnCount = _columnCount;
		if (columnCount == 0) {
			return const <_Cell>[];
		}
		final cells = <_Cell>[];
		for (final tc in _tbl.iterTcs()) {
			for (var spanIdx = 0; spanIdx < tc.gridSpan; spanIdx++) {
				if (tc.vMerge == ST_Merge.CONTINUE) {
					if (cells.length >= columnCount) {
						cells.add(cells[cells.length - columnCount]);
					} else {
						cells.add(_Cell(tc, this));
					}
					continue;
				}
				if (spanIdx > 0 && cells.isNotEmpty) {
					cells.add(cells.last);
					continue;
				}
				cells.add(_Cell(tc, this));
			}
		}
		return cells;
	}

	int get _columnCount => _tbl.colCount;
	int get _rowCount => _tbl.trList.length;
	CT_TblPr get _tblPr => _tbl.tblPr;

	void _ensureValidColumn(int columnIdx) {
		if (_columnCount == 0) {
			throw RangeError('table has no columns');
		}
		final maxIdx = _columnCount - 1;
		if (columnIdx < 0 || columnIdx > maxIdx) {
			throw RangeError.range(columnIdx, 0, maxIdx, 'columnIdx');
		}
	}

	void _ensureValidRow(int rowIdx) {
		if (_rowCount == 0) {
			throw RangeError('table has no rows');
		}
		final maxIdx = _rowCount - 1;
		if (rowIdx < 0 || rowIdx > maxIdx) {
			throw RangeError.range(rowIdx, 0, maxIdx, 'rowIdx');
		}
	}
}

class _Cell extends BlockItemContainer {
	_Cell(this._tc, Table table)
			: _table = table,
				super(_tc, table);

	final CT_Tc _tc;
	final Table _table;

	Table addTable(int rows, int cols, [Length? width]) {
		final tableWidth = width ?? _tc.width ?? Inches(1);
		final table = super.addTable(rows, cols, tableWidth);
		addParagraph();
		return table;
	}

	int get gridSpan => _tc.gridSpan;

	_Cell merge(_Cell other) {
		final merged = _tc.merge(other._tc);
		return _Cell(merged, _table);
	}

	String get text => super.paragraphs.map((p) => p.text).join('\n');
	set text(String value) {
		_tc.clearContent();
		final paragraph = _tc.addP();
		final run = paragraph.addR();
		run.text = value;
	}

	WD_CELL_VERTICAL_ALIGNMENT? get verticalAlignment => _tc.tcPr?.vAlignVal;
	set verticalAlignment(WD_CELL_VERTICAL_ALIGNMENT? value) {
		final tcPr = _tc.getOrAddTcPr();
		tcPr.vAlignVal = value;
	}

	Length? get width => _tc.width;
	set width(Length? value) => _tc.width = value;

	Table get table => _table;
}

class _Column {
	_Column(this._gridCol, this._table);

	final CT_TblGridCol _gridCol;
	final Table _table;

	List<_Cell> get cells => List<_Cell>.unmodifiable(_table.columnCells(_index));
	Table get table => _table;

	Length? get width => _gridCol.w;
	set width(Length? value) => _gridCol.w = value;

	int get _index => _gridCol.gridColIdx;
}

class _Columns extends IterableBase<_Column> {
	_Columns(this._tbl, this._table);

	final CT_Tbl _tbl;
	final Table _table;

	@override
	Iterator<_Column> get iterator =>
			_tbl.tblGrid.gridColList.map((gridCol) => _Column(gridCol, _table)).iterator;

	_Column operator [](int index) {
		final cols = _tbl.tblGrid.gridColList;
		if (index < 0 || index >= cols.length) {
			throw RangeError.index(index, cols, 'index');
		}
		return _Column(cols[index], _table);
	}

	int get length => _tbl.tblGrid.gridColList.length;
	Table get table => _table;
}

class _Row {
	_Row(this._tr, this._table);

	final CT_Row _tr;
	final Table _table;

	List<_Cell> get cells {
		final rowCells = <_Cell>[];
		for (final tc in _tr.tcList) {
			rowCells.addAll(_expandTc(tc));
		}
		return List<_Cell>.unmodifiable(rowCells);
	}

	int get gridColsAfter => _tr.gridAfter;
	int get gridColsBefore => _tr.gridBefore;

	Length? get height => _tr.trHeight_val;
	set height(Length? value) => _tr.trHeight_val = value;

	WD_ROW_HEIGHT_RULE? get heightRule => _tr.trHeight_hRule;
	set heightRule(WD_ROW_HEIGHT_RULE? value) => _tr.trHeight_hRule = value;

	Table get table => _table;

	Iterable<_Cell> _expandTc(CT_Tc tc) sync* {
		if (tc.vMerge == ST_Merge.CONTINUE) {
			yield* _expandTc(tc.tcAbove);
			return;
		}
		final cell = _Cell(tc, _table);
		for (var i = 0; i < tc.gridSpan; i++) {
			yield cell;
		}
	}
}

class _Rows extends IterableBase<_Row> {
	_Rows(this._tbl, this._table);

	final CT_Tbl _tbl;
	final Table _table;

	@override
	Iterator<_Row> get iterator =>
			_tbl.trList.map((tr) => _Row(tr, _table)).iterator;

	_Row operator [](int index) {
		final rows = _tbl.trList;
		if (index < 0 || index >= rows.length) {
			throw RangeError.index(index, rows, 'index');
		}
		return _Row(rows[index], _table);
	}

	int get length => _tbl.trList.length;
	Table get table => _table;
}
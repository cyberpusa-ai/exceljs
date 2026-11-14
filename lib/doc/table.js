/* eslint-disable max-classes-per-file */
const colCache = require('../utils/col-cache');

// Excel's rules for table names: start with a letter, '_' or '\\', then letters, digits, '.' or
// '_'; at most 255 characters; and the name must not be read as a cell reference.
// Any non-ASCII character is treated as a letter so that names Excel accepts keep working.
const tableNameRegex = /^[A-Za-z_\\\u0080-\uFFFF][A-Za-z0-9._\u0080-\uFFFF]*$/;

function isCellReference(name) {
  if (/^[RC]$/i.test(name) || /^R\d*C\d*$/i.test(name)) {
    // 'R', 'C' and R1C1-style references such as 'R1C1' or 'RC'
    return true;
  }
  const match = /^([A-Z]{1,3})(\d+)$/i.exec(name);
  if (!match) {
    return false;
  }
  // letters past XFD are not a column, so a name like 'ZZZ1' is fine
  const letters = match[1].toUpperCase();
  const isColumn = letters.length < 3 || letters <= 'XFD';
  const row = parseInt(match[2], 10);
  return isColumn && row >= 1 && row <= 1048576;
}

function isValidTableName(name) {
  return (
    typeof name === 'string' &&
    name.length <= 255 &&
    tableNameRegex.test(name) &&
    !isCellReference(name)
  );
}

class Column {
  // wrapper around column model, allowing access and manipulation
  constructor(table, column, index) {
    this.table = table;
    this.column = column;
    this.index = index;
  }

  _set(name, value) {
    this.table.cacheState();
    this.column[name] = value;
  }

  /* eslint-disable lines-between-class-members */
  get name() {
    return this.column.name;
  }
  set name(value) {
    this._set('name', value);
  }

  get filterButton() {
    return this.column.filterButton;
  }
  set filterButton(value) {
    this.column.filterButton = value;
  }

  get style() {
    return this.column.style;
  }
  set style(value) {
    this.column.style = value;
  }

  get totalsRowLabel() {
    return this.column.totalsRowLabel;
  }
  set totalsRowLabel(value) {
    this._set('totalsRowLabel', value);
  }

  get totalsRowFunction() {
    return this.column.totalsRowFunction;
  }
  set totalsRowFunction(value) {
    this._set('totalsRowFunction', value);
  }

  get totalsRowResult() {
    return this.column.totalsRowResult;
  }
  set totalsRowResult(value) {
    this._set('totalsRowResult', value);
  }

  get totalsRowFormula() {
    return this.column.totalsRowFormula;
  }
  set totalsRowFormula(value) {
    this._set('totalsRowFormula', value);
  }
  /* eslint-enable lines-between-class-members */
}

class Table {
  constructor(worksheet, table) {
    this.worksheet = worksheet;
    if (table) {
      this.table = table;
      // check things are ok first
      this.validate();

      this.store();
    }
  }

  getFormula(column) {
    // get the correct formula to apply to the totals row
    switch (column.totalsRowFunction) {
      case 'none':
        return null;
      case 'average':
        return `SUBTOTAL(101,${this.table.name}[${column.name}])`;
      case 'countNums':
        return `SUBTOTAL(102,${this.table.name}[${column.name}])`;
      case 'count':
        return `SUBTOTAL(103,${this.table.name}[${column.name}])`;
      case 'max':
        return `SUBTOTAL(104,${this.table.name}[${column.name}])`;
      case 'min':
        return `SUBTOTAL(105,${this.table.name}[${column.name}])`;
      case 'stdDev':
        return `SUBTOTAL(106,${this.table.name}[${column.name}])`;
      case 'var':
        return `SUBTOTAL(107,${this.table.name}[${column.name}])`;
      case 'sum':
        return `SUBTOTAL(109,${this.table.name}[${column.name}])`;
      case 'custom':
        return column.totalsRowFormula;
      default:
        throw new Error(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
    }
  }

  get width() {
    // width of the table
    return this.table.columns.length;
  }

  get height() {
    // height of the table data. Excel needs at least one data row, so an empty table
    // keeps one blank row on the sheet (it is not added to table.rows)
    return Math.max(this.table.rows.length, 1);
  }

  get filterHeight() {
    // height of the table data plus optional header row
    return this.height + (this.table.headerRow ? 1 : 0);
  }

  get tableHeight() {
    // full height of the table on the sheet
    return this.filterHeight + (this.table.totalsRow ? 1 : 0);
  }

  validate() {
    const {table} = this;
    // set defaults and check is valid
    const assign = (o, name, dflt) => {
      if (o[name] === undefined) {
        o[name] = dflt;
      }
    };
    assign(table, 'headerRow', true);
    assign(table, 'totalsRow', false);

    assign(table, 'style', {});
    assign(table.style, 'theme', 'TableStyleMedium2');
    assign(table.style, 'showFirstColumn', false);
    assign(table.style, 'showLastColumn', false);
    assign(table.style, 'showRowStripes', false);
    assign(table.style, 'showColumnStripes', false);

    const assert = (test, message) => {
      if (!test) {
        throw new Error(message);
      }
    };
    assert(table.name, 'Table must have a name');
    assert(
      isValidTableName(table.name),
      `Invalid table name '${table.name}'. A table name must start with a letter, underscore ` +
        'or backslash, contain only letters, digits, periods and underscores, ' +
        'not look like a cell reference and be at most 255 characters long.'
    );
    assert(table.ref, 'Table must have ref');
    assert(table.columns, 'Table must have column definitions');
    assert(table.rows, 'Table must have row definitions');

    table.tl = colCache.decodeAddress(table.ref);
    const {row, col} = table.tl;
    assert(row > 0, 'Table must be on valid row');
    assert(col > 0, 'Table must be on valid col');

    const {width, filterHeight, tableHeight} = this;

    // autoFilterRef is a range that includes optional headers only
    table.autoFilterRef = colCache.encode(row, col, row + filterHeight - 1, col + width - 1);

    // tableRef is a range that includes optional headers and totals
    table.tableRef = colCache.encode(row, col, row + tableHeight - 1, col + width - 1);

    table.columns.forEach((column, i) => {
      assert(column.name, `Column ${i} must have a name`);
      if (i === 0) {
        assign(column, 'totalsRowLabel', 'Total');
      } else {
        assign(column, 'totalsRowFunction', 'none');
        column.totalsRowFormula = this.getFormula(column);
      }
    });
  }

  store() {
    // where the table needs to store table data, headers, footers in
    // the sheet...
    const assignStyle = (cell, style) => {
      if (style) {
        Object.keys(style).forEach(key => {
          cell.style[key] = style[key];
        });
      }
    };

    const {worksheet, table} = this;
    const {row, col} = table.tl;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const {style, name} = column;
        const cell = r.getCell(col + j);
        cell.value = name;
        assignStyle(cell, style);
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value;

        assignStyle(cell, table.columns[j].style);
      });
    });
    if (!table.rows.length) {
      // blank placeholder row so the table keeps a valid ref
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        cell.value = null;
        assignStyle(cell, column.style);
      });
    }

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula,
              result: column.totalsRowResult,
            };
          } else {
            cell.value = null;
          }
        }

        assignStyle(cell, column.style);
      });
    }
  }

  load(worksheet) {
    // where the table will read necessary features from a loaded sheet
    const {table} = this;
    const {row, col} = table.tl;
    let count = 0;
    if (table.headerRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        cell.value = column.name;
      });
    }
    table.rows.forEach(data => {
      const r = worksheet.getRow(row + count++);
      data.forEach((value, j) => {
        const cell = r.getCell(col + j);
        cell.value = value;
      });
    });
    if (!table.rows.length) {
      // skip the blank placeholder row of an empty table
      count++;
    }

    if (table.totalsRow) {
      const r = worksheet.getRow(row + count++);
      table.columns.forEach((column, j) => {
        const cell = r.getCell(col + j);
        if (j === 0) {
          cell.value = column.totalsRowLabel;
        } else {
          const formula = this.getFormula(column);
          if (formula) {
            cell.value = {
              formula: column.totalsRowFormula,
              result: column.totalsRowResult,
            };
          }
        }
      });
    }
  }

  get model() {
    return this.table;
  }

  set model(value) {
    this.table = value;
    if (this.worksheet && !value.rows) {
      // a table parsed from a file has no rows: read them from the sheet cells
      this.loadRows();
    }
  }

  loadRows() {
    // rebuild ref, column names, rows and totals row formulas of a table parsed from a file
    // from the cells inside its tableRef. This only reads the sheet, it never writes to it.
    const {worksheet, table} = this;
    if (table.headerRow === undefined) {
      table.headerRow = true;
    }
    if (table.totalsRow === undefined) {
      table.totalsRow = false;
    }
    if (!table.style) {
      table.style = {};
    }
    table.columns = table.columns || [];
    const range = colCache.decode(table.tableRef || table.ref);
    const top = range.top || range.row;
    const left = range.left || range.col;
    const bottom = range.bottom || top;
    if (!table.ref) {
      table.ref = colCache.encodeAddress(top, left);
    }
    const cellValue = (r, c) => {
      const cell = worksheet.findCell(r, c);
      return cell ? cell.value : null;
    };

    let first = top;
    let last = bottom;
    if (table.headerRow) {
      table.columns.forEach((column, j) => {
        if (!column.name) {
          const value = cellValue(top, left + j);
          column.name = value === null || value === undefined ? `Column${j + 1}` : `${value}`;
        }
      });
      first++;
    }
    if (table.totalsRow) {
      table.columns.forEach((column, j) => {
        const cell = worksheet.findCell(bottom, left + j);
        if (!cell) {
          return;
        }
        if (j === 0) {
          if (column.totalsRowLabel === undefined && typeof cell.value === 'string') {
            column.totalsRowLabel = cell.value;
          }
        } else if (cell.formula) {
          if (column.totalsRowFunction === 'custom' && column.totalsRowFormula === undefined) {
            column.totalsRowFormula = cell.formula;
          }
          if (column.totalsRowResult === undefined) {
            column.totalsRowResult = cell.result;
          }
        }
      });
      last--;
    }

    table.rows = [];
    for (let r = first; r <= last; r++) {
      table.rows.push(table.columns.map((column, j) => cellValue(r, left + j)));
    }
    if (table.rows.length === 1 && table.rows[0].every(v => v === null || v === undefined)) {
      // the blank placeholder row of an empty table
      table.rows = [];
    }
  }

  // ================================================================
  // TODO: Mutating methods
  cacheState() {
    if (!this._cache) {
      this._cache = {
        ref: this.ref,
        width: this.width,
        tableHeight: this.tableHeight,
      };
    }
  }

  commit() {
    // changes may have been made that might have on-sheet effects
    if (!this._cache) {
      return;
    }

    // check things are ok first
    this.validate();

    const ref = colCache.decodeAddress(this._cache.ref);
    if (this.ref !== this._cache.ref) {
      // wipe out whole table footprint at previous location
      for (let i = 0; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }
    } else {
      // clear out below table if it has shrunk
      for (let i = this.tableHeight; i < this._cache.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = 0; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }

      // clear out to right of table if it has lost columns
      for (let i = 0; i < this.tableHeight; i++) {
        const row = this.worksheet.getRow(ref.row + i);
        for (let j = this.width; j < this._cache.width; j++) {
          const cell = row.getCell(ref.col + j);
          cell.value = null;
        }
      }
    }

    this.store();
  }

  addRow(values, rowNumber) {
    // Add a row of data, either insert at rowNumber or append
    this.cacheState();

    if (rowNumber === undefined) {
      this.table.rows.push(values);
    } else {
      this.table.rows.splice(rowNumber, 0, values);
    }
  }

  removeRows(rowIndex, count = 1) {
    // Remove a rows of data
    this.cacheState();
    this.table.rows.splice(rowIndex, count);
  }

  getColumn(colIndex) {
    const column = this.table.columns[colIndex];
    return new Column(this, column, colIndex);
  }

  addColumn(column, values, colIndex) {
    // Add a new column, including column defn and values
    // Inserts at colNumber or adds to the right
    this.cacheState();

    if (colIndex === undefined) {
      this.table.columns.push(column);
      this.table.rows.forEach((row, i) => {
        row.push(values[i]);
      });
    } else {
      this.table.columns.splice(colIndex, 0, column);
      this.table.rows.forEach((row, i) => {
        row.splice(colIndex, 0, values[i]);
      });
    }
  }

  removeColumns(colIndex, count = 1) {
    // Remove a column with data
    this.cacheState();

    this.table.columns.splice(colIndex, count);
    this.table.rows.forEach(row => {
      row.splice(colIndex, count);
    });
  }

  _assign(target, prop, value) {
    this.cacheState();
    target[prop] = value;
  }

  /* eslint-disable lines-between-class-members */
  get ref() {
    return this.table.ref;
  }
  set ref(value) {
    this._assign(this.table, 'ref', value);
  }

  get name() {
    return this.table.name;
  }
  set name(value) {
    this.table.name = value;
  }

  get displayName() {
    return this.table.displyName || this.table.name;
  }
  set displayNamename(value) {
    this.table.displayName = value;
  }

  get headerRow() {
    return this.table.headerRow;
  }
  set headerRow(value) {
    this._assign(this.table, 'headerRow', value);
  }

  get totalsRow() {
    return this.table.totalsRow;
  }
  set totalsRow(value) {
    this._assign(this.table, 'totalsRow', value);
  }

  get theme() {
    return this.table.style.name;
  }
  set theme(value) {
    this.table.style.name = value;
  }

  get showFirstColumn() {
    return this.table.style.showFirstColumn;
  }
  set showFirstColumn(value) {
    this.table.style.showFirstColumn = value;
  }

  get showLastColumn() {
    return this.table.style.showLastColumn;
  }
  set showLastColumn(value) {
    this.table.style.showLastColumn = value;
  }

  get showRowStripes() {
    return this.table.style.showRowStripes;
  }
  set showRowStripes(value) {
    this.table.style.showRowStripes = value;
  }

  get showColumnStripes() {
    return this.table.style.showColumnStripes;
  }
  set showColumnStripes(value) {
    this.table.style.showColumnStripes = value;
  }
  /* eslint-enable lines-between-class-members */
}

module.exports = Table;

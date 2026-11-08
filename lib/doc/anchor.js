'use strict';

const colCache = require('../utils/col-cache');

// Offsets in drawing anchors are in EMUs. Excel renders at 96 DPI (9525 EMU per pixel), a column
// width unit is the maximum digit width of the default font (7px for Calibri 11) and row heights
// are in points (12700 EMU per point).
const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;
const MAX_DIGIT_WIDTH_PX = 7;
const DEFAULT_COL_WIDTH_EMU = 64 * EMU_PER_PIXEL; // Excel's default column is 64px wide
const DEFAULT_ROW_HEIGHT_PT = 15;

class Anchor {
  constructor(worksheet, address, offset = 0) {
    this.worksheet = worksheet;

    if (!address) {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    } else if (typeof address === 'string') {
      const decoded = colCache.decodeAddress(address);
      this.nativeCol = decoded.col + offset;
      this.nativeColOff = 0;
      this.nativeRow = decoded.row + offset;
      this.nativeRowOff = 0;
    } else if (address.nativeCol !== undefined) {
      this.nativeCol = address.nativeCol || 0;
      this.nativeColOff = address.nativeColOff || 0;
      this.nativeRow = address.nativeRow || 0;
      this.nativeRowOff = address.nativeRowOff || 0;
    } else if (address.col !== undefined) {
      this.col = address.col + offset;
      this.row = address.row + offset;
    } else {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    }
  }

  static asInstance(model) {
    return model instanceof Anchor || model == null ? model : new Anchor(model);
  }

  get col() {
    return this.nativeCol + (Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth);
  }

  set col(v) {
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
  }

  get row() {
    return this.nativeRow + (Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight);
  }

  set row(v) {
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
  }

  get colWidth() {
    const column = this._getColumn();
    const properties = this.worksheet && this.worksheet.properties;
    const width = (column && column.width) || (properties && properties.defaultColWidth);
    return width ? Math.floor(width * MAX_DIGIT_WIDTH_PX * EMU_PER_PIXEL) : DEFAULT_COL_WIDTH_EMU;
  }

  get rowHeight() {
    const row = this._getRow();
    const properties = this.worksheet && this.worksheet.properties;
    const height =
      (row && row.height) || (properties && properties.defaultRowHeight) || DEFAULT_ROW_HEIGHT_PT;
    return Math.floor(height * EMU_PER_POINT);
  }

  _getColumn() {
    return this.worksheet && typeof this.worksheet.getColumn === 'function'
      ? this.worksheet.getColumn(this.nativeCol + 1)
      : undefined;
  }

  _getRow() {
    return this.worksheet && typeof this.worksheet.getRow === 'function'
      ? this.worksheet.getRow(this.nativeRow + 1)
      : undefined;
  }

  get model() {
    return {
      nativeCol: this.nativeCol,
      nativeColOff: this.nativeColOff,
      nativeRow: this.nativeRow,
      nativeRowOff: this.nativeRowOff,
    };
  }

  set model(value) {
    this.nativeCol = value.nativeCol;
    this.nativeColOff = value.nativeColOff;
    this.nativeRow = value.nativeRow;
    this.nativeRowOff = value.nativeRowOff;
  }
}

module.exports = Anchor;

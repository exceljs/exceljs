'use strict';

const colCache = require('../utils/col-cache');

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
    let width = this.worksheet.properties.defaultColWidth != undefined ? this.worksheet.properties.defaultColWidth : 9.14285714285714;
    return this.worksheet &&
      this.worksheet.getColumn(this.nativeCol + 1) &&
      this.worksheet.getColumn(this.nativeCol + 1).isCustomWidth
      ? Math.floor(this.worksheet.getColumn(this.nativeCol + 1).width/width * 640000)
      : 640000;
  }

  get rowHeight() {
    let height = this.worksheet.properties.defaultRowHeight != undefined ? this.worksheet.properties.defaultRowHeight : 15;
    return this.worksheet &&
      this.worksheet.getRow(this.nativeRow + 1) &&
      this.worksheet.getRow(this.nativeRow + 1).height
      ? Math.floor(this.worksheet.getRow(this.nativeRow + 1).height/height * 180000)
      : 180000;
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

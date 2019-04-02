'use strict';

var Anchor = module.exports = function (model = {}) {
  this.nativeCol = model.nativeCol || 0;
  this.nativeColOff = model.nativeColOff || 0;
  this.nativeRow = model.nativeRow || 0;
  this.nativeRowOff = model.nativeRowOff || 0;

  if (model.col) {
    this.col = model.col;
  }

  if (model.row) {
    this.row = model.row;
  }
};

Anchor.asInstance = function (model) {
  return model instanceof Anchor || model == null ? model : new Anchor(model);
};

Anchor.prototype = {
  worksheet: null,
  get col() {
    return this.nativeCol + Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth;
  },
  set col(v) {
    if (v === this.col) return;
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
  },
  get row() {
    return this.nativeRow + Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight;
  },
  set row(v) {
    if (v === this.row) return;
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
  },
  get colWidth() {
    return this.worksheet && this.worksheet.getColumn(this.nativeCol + 1) && this.worksheet.getColumn(this.nativeCol + 1).isCustomWidth ?
      Math.floor(this.worksheet.getColumn(this.nativeCol + 1).width * 10000) :
      640000;
  },
  get rowHeight() {
    return this.worksheet && this.worksheet.getRow(this.nativeRow + 1) && this.worksheet.getRow(this.nativeRow + 1).height ?
      Math.floor(this.worksheet.getRow(this.nativeRow + 1).height * 10000) :
      180000;
  }
};

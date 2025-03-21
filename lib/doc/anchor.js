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
    return this.worksheet &&
      this.worksheet.getColumn(this.nativeCol + 1) &&
      this.worksheet.getColumn(this.nativeCol + 1).isCustomWidth
      ? Math.floor(this.worksheet.getColumn(this.nativeCol + 1).width * 10000)
      : 640000;
  }

  get rowHeight() {
    return this.worksheet &&
      this.worksheet.getRow(this.nativeRow + 1) &&
      this.worksheet.getRow(this.nativeRow + 1).height
      ? Math.floor(this.worksheet.getRow(this.nativeRow + 1).height * 10000)
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

  // Converts Anchor back into cell reference string (e.g., "A1")
  toString() {
    const colStr = colCache.encodeCol(this.nativeCol);
    const rowStr = this.nativeRow + 1;
    return `${colStr}${rowStr}`;
  }

  // Compares two Anchor instances for equality
  equals(otherAnchor) {
    return otherAnchor instanceof Anchor &&
      this.nativeCol === otherAnchor.nativeCol &&
      this.nativeRow === otherAnchor.nativeRow;
  }

  // Validate if a cell address is valid (e.g., "A1", "Z100")
  static isValidAddress(address) {
    const regex = /^[A-Z]+[0-9]+$/;
    return regex.test(address);
  }

  // Range Support: Returns the range as a string (e.g., "A1:B2")
  static createRange(start, end) {
    if (start instanceof Anchor && end instanceof Anchor) {
      return `${start.toString()}:${end.toString()}`;
    }
    throw new Error("Both start and end must be instances of Anchor.");
  }

  // Example of an extension: Allow shifting by a certain number of rows and columns
  shift(offsetRow, offsetCol) {
    this.nativeRow += offsetRow;
    this.nativeCol += offsetCol;
  }

  // Check if the anchor is within the valid range of the worksheet
  isValid() {
    const maxRow = this.worksheet.rowCount;
    const maxCol = this.worksheet.columnCount;
    return this.nativeRow >= 0 && this.nativeRow < maxRow &&
      this.nativeCol >= 0 && this.nativeCol < maxCol;
  }

  // Get the worksheet context (address, sheet name)
  getContext() {
    return {
      worksheet: this.worksheet.name,
      address: this.toString(),
    };
  }

  // Get relative address (e.g., distance between two anchors in terms of rows/cols)
  getRelativeAddress(otherAnchor) {
    if (otherAnchor instanceof Anchor) {
      return {
        rowDiff: this.nativeRow - otherAnchor.nativeRow,
        colDiff: this.nativeCol - otherAnchor.nativeCol,
      };
    }
    throw new Error("Argument must be an instance of Anchor.");
  }

  // Create a range that spans multiple cells (e.g., "A1:C3")
  static createMultiCellRange(start, end) {
    if (start instanceof Anchor && end instanceof Anchor) {
      const startCol = colCache.encodeCol(start.nativeCol);
      const startRow = start.nativeRow + 1;
      const endCol = colCache.encodeCol(end.nativeCol);
      const endRow = end.nativeRow + 1;
      return `${startCol}${startRow}:${endCol}${endRow}`;
    }
    throw new Error("Both start and end must be instances of Anchor.");
  }

  // Clone the current anchor to create a copy
  clone() {
    return new Anchor(this.worksheet, this.toString());
  }
}

module.exports = Anchor;

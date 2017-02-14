/**
 * Copyright (c) 2014 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var _ = require('../utils/under-dash');

var Enums = require('./enums');
var colCache = require('./../utils/col-cache');
var Cell = require('./cell');

var Row = module.exports = function (worksheet, number) {
  this._worksheet = worksheet;
  this._number = number;
  this._cells = [];
  this.style = {};
  this.outlineLevel = 0;
};

Row.prototype = {
  // return the row number
  get number() {
    return this._number;
  },

  get worksheet() {
    return this._worksheet;
  },

  // Inform Streaming Writer that this row (and all rows before it) are complete
  // and ready to write. Has no effect on Worksheet document
  commit: function () {
    this._worksheet._commitRow(this);
  },

  // helps GC by breaking cyclic references
  destroy: function () {
    delete this._worksheet;
    delete this._cells;
    delete this.style;
  },

  findCell: function (colNumber) {
    return this._cells[colNumber - 1];
  },

  // given {address, row, col}, find or create new cell
  _getCell: function (address) {
    var cell = this._cells[address.col - 1];
    if (!cell) {
      var column = this._worksheet.getColumn(address.col);
      cell = new Cell(this, column, address.address);
      this._cells[address.col - 1] = cell;
    }
    return cell;
  },

  // get cell by key, letter or column number
  getCell: function (col) {
    if (typeof col == 'string') {
      // is it a key?
      var column = this._worksheet._keys[col];
      if (column) {
        col = column.number;
      } else {
        col = colCache.l2n(col);
      }
    }
    return this._cells[col - 1] ||
      this._getCell({
        address: colCache.encodeAddress(this._number, col),
        row: this._number,
        col: col
      });
  },

  // remove cell(s) and shift all higher cells down by count
  splice: function(start, count) {
    var inserts = Array.prototype.slice.call(arguments, 2);
    var nKeep = start + count;
    var nExpand = inserts.length - count;
    var nEnd = this._cells.length;
    var i, cSrc, cDst;
    if (nExpand < 0) {
      // remove cells
      for (i = nKeep; i <= nEnd; i++) {
        cSrc = this._cells[i-1];
        if (cSrc) {
          this.getCell(i + nExpand).value = cSrc.value;
          cSrc.value = null;
        } else {
          cDst = this._cells[i + nExpand - 1];
          if (cDst) {
            cDst.value = null;
          }
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        cSrc = this._cells[i-1];
        if (cSrc) {
          this.getCell(i + nExpand).value = cSrc.value;
        } else {
          this._cells[i + nExpand - 1] = undefined;
        }
      }
    }

    // now add the new values
    for (i = 0; i < inserts.length; i++) {
      this.getCell(start + i).value = inserts[i];
    }
  },

  // Iterate over all non-null cells in this row
  eachCell: function (options, iteratee) {
    if (!iteratee) {
      iteratee = options;
      options = null;
    }
    if (options && options.includeEmpty) {
      var n = this._cells.length;
      for (var i = 1; i <= n; i++) {
        iteratee(this.getCell(i), i);
      }
    } else {
      this._cells.forEach(function (cell, index) {
        if (cell && (cell.type != Enums.ValueType.Null)) {
          iteratee(cell, index + 1);
        }
      });
    }
  },

  // return a sparse array of cell values
  get values() {
    var values = [];
    this._cells.forEach(function (cell) {
      if (cell && (cell.type != Enums.ValueType.Null)) {
        values[cell.col] = cell.value;
      }
    });
    return values;
  },

  // set the values by contiguous or sparse array, or by key'd object literal
  set values(value) {
    var self = this;

    // this operation is not additive - any prior cells are removed
    this._cells = [];
    if (!value) {
      // empty row
    } else if (value instanceof Array) {
      var offset = 0;
      if (value.hasOwnProperty('0')) {
        // contiguous array - start at column 1
        offset = 1;
      }
      value.forEach(function (item, index) {
        if (item !== undefined) {
          self._getCell({
            address: colCache.encodeAddress(self._number, index + offset),
            row: self._number,
            col: index + offset
          }).value = item;
        }
      });
    } else {
      // assume object with column keys
      _.each(this._worksheet._keys, function (column, key) {
        if (value[key] !== undefined) {
          self._getCell({
            address: colCache.encodeAddress(self._number, column.number),
            row: self._number,
            col: column.number
          }).value = value[key];
        }
      });
    }
  },

  // returns true if the row includes at least one cell with a value
  get hasValues() {
    return _.some(this._cells, function (cell) {
      return cell && (cell.type != Enums.ValueType.Null);
    });
  },

  get cellCount() {
    return this._cells.length;
  },
  get actualCellCount() {
    var count = 0;
    this.eachCell(function() {
      count++;
    });
    return count;
  },

  // get the min and max column number for the non-null cells in this row or null
  get dimensions() {
    var min = 0;
    var max = 0;
    this._cells.forEach(function (cell) {
      if (cell && (cell.type != Enums.ValueType.Null)) {
        if (!min || (min > cell.col)) {
          min = cell.col;
        }
        if (max < cell.col) {
          max = cell.col;
        }
      }
    });
    return min > 0 ? {
      min: min,
      max: max
    } : null;
  },

  // =========================================================================
  // styles
  _applyStyle: function (name, value) {
    this.style[name] = value;
    this._cells.forEach(function (cell) {
      if (cell) { cell[name] = value; }
    });
    return value;
  },

  get numFmt() {
    return this.style.numFmt;
  },
  set numFmt(value) {
    return this._applyStyle('numFmt', value);
  },
  get font() {
    return this.style.font;
  },
  set font(value) {
    return this._applyStyle('font', value);
  },
  get alignment() {
    return this.style.alignment;
  },
  set alignment(value) {
    return this._applyStyle('alignment', value);
  },
  get border() {
    return this.style.border;
  },
  set border(value) {
    return this._applyStyle('border', value);
  },
  get fill() {
    return this.style.fill;
  },
  set fill(value) {
    return this._applyStyle('fill', value);
  },

  get hidden() {
    return !!this._hidden;
  },
  set hidden(value) {
    return this._hidden = value;
  },

  get outlineLevel() {
    return this._outlineLevel || 0;
  },
  set outlineLevel(value) {
    this._outlineLevel = value;
    return value;
  },
  get collapsed() {
    return !!(this._outlineLevel && (this._outlineLevel >= this._worksheet.properties.outlineLevelRow));
  },

  // =========================================================================
  get model() {
    var cells = [];
    var min = 0;
    var max = 0;
    this._cells.forEach(function (cell) {
      if (cell) {
        var cellModel = cell.model;
        if (cellModel) {
          if (!min || (min > cell.col)) {
            min = cell.col;
          }
          if (max < cell.col) {
            max = cell.col;
          }
          cells.push(cellModel);
        }
      }
    });

    return (this.height || cells.length) ? {
      cells: cells,
      number: this.number,
      min: min,
      max: max,
      height: this.height,
      style: this.style,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
      collapsed: this.collapsed
    } : null;
  },
  set model(value) {
    if (value.number != this._number) {
      throw new Error('Invalid row number in model');
    }
    var self = this;
    this._cells = [];
    value.cells.forEach(function (cellModel) {
      switch (cellModel.type) {
        case Cell.Types.Merge:
          // special case - don't add this types
          break;
        default:
          var cell = self._getCell(colCache.decodeAddress(cellModel.address));
          cell.model = cellModel;
          break;
      }
    });

    if (value.height) {
      this.height = value.height;
    } else {
      delete this.height;
    }

    this.hidden = value.hidden;
    this.outlineLevel = value.outlineLevel || 0;

    this.style = value.style || {};
  }
};
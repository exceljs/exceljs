/**
 * Copyright (c) 2014 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software'), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var _ = require('underscore');

var utils = require('./../utils/utils');
var colCache = require('./../utils/col-cache');
var Range = require('./range');
var Row = require('./row');
var Column = require('./column');
var Enums = require('./enums');
var DataValidations = require('./data-validations');
var WorksheetProperties = require('./worksheet-properties');

// Worksheet requirements
//  Operate as sheet inside workbook or standalone
//  Load and Save from file and stream
//  Access/Add/Delete individual cells
//  Manage column widths and row heights

var Worksheet = module.exports = function(options) {
  options = options || {};

  // in a workbook, each sheet will have a number
  this.id = options.id;

  // and a name
  this.name = options.name || 'Sheet' + this.id;

  // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
  // Note: _rows is zero based. Must subtract 1 to go from cell.row to index
  this._rows = [];

  // column definitions
  this._columns = null;

  // column keys (addRow convenience): key ==> this._collumns index
  this._keys = {};

  // keep record of all merges
  this._merges = {};

  this._workbook = options.workbook;

  // for tabColor, default row height, outline levels, etc
  this.properties = new WorksheetProperties(this, options.properties);
  
  this.dataValidations = new DataValidations();
};

Worksheet.prototype = {

  get workbook() {
    return this._workbook;
  },

  // when you're done with this worksheet, call this to remove from workbook
  destroy: function() {
    this._workbook._removeWorksheet(this);
  },

  // Get the bounding range of the cells in this worksheet
  get dimensions() {
    var dimensions = new Range();
    _.each(this._rows, function(row) {
      var rowDims = row.dimensions;
      if (rowDims) {
        dimensions.expand(row.number, rowDims.min, row.number, rowDims.max);
      }
    });
    return dimensions;
  },

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns() {
    return this._columns;
  },

  // set the columns from an array of column definitions.
  // Note: any headers defined will overwrite existing values.
  set columns(value) {
    var self = this;

    // calculate max header row count
    this._headerRowCount = value.reduce(function(pv,cv) {
      var headerCount = cv.header ? 1 : (cv.headers ? cv.headers.length : 0);
      return Math.max(pv, headerCount)
    }, 0);

    // construct Column objects
    var count = 1;
    var _columns = this._columns = [];
    _.each(value, function(defn) {
      var column = new Column(self, count++, false);
      _columns.push(column);
      column.defn = defn;
    });
  },

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn: function(c) {
    if (typeof c == 'string'){
      // if it matches a key'd column, return that
      var col = this._keys[c];
      if (col) return col;

      // otherise, assume letter
      c = colCache.l2n(c);
    }
    if (!this._columns) { this._columns = []; }
    if (c > this._columns.length) {
      var n = this._columns.length + 1;
      while (n <= c) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[c-1];
  },

  // =========================================================================
  // Rows

  _commitRow: function() {
    // nop - allows streaming reader to fill a document
  },

  get _nextRow() {
    return this._rows.length + 1;
  },

  get lastRow() {
    if (this._rows.length) {
      return this._rows[this._rows.length-1];
    } else {
      return undefined;
    }
  },

  // find a row (if exists) by row number
  findRow: function(r) {
    return this._rows[r-1];
  },

  // get a row by row number.
  getRow: function(r) {
    var row = this._rows[r-1];
    if (!row) {
      row = this._rows[r-1] = new Row(this, r);
    }
    return row;
  },
  addRow: function(value) {
    var row = this.getRow(this._nextRow);
    row.values = value;
    return row;
  },
  addRows: function(value) {
    var self = this;
    value.forEach(function(row) {
      self.addRow(row);
    });
  },

  // iterate over every row in the worksheet, including maybe empty rows
  eachRow: function(options, iteratee) {
    if (!iteratee) {
      iteratee = options;
      options = undefined;
    }
    if (options && options.includeEmpty) {
      var n = this._rows.length;
      for (var i = 1; i <= n; i++) {
        iteratee(this.getRow(i), i);
      }
    } else {
      this._rows.forEach(function(row) {
        if (row.hasValues) {
          iteratee(row, row.number);
        }
      });
    }
  },

  // return all rows as sparse array
  getSheetValues: function() {
    var rows = [];
    this._rows.forEach(function(row) {
      rows[row.number] = row.values;
    });
    return rows;
  },

  // =========================================================================
  // Cells

  // returns the cell at [r,c] or address given by r. If not found, return undefined
  findCell: function(r, c) {
    var address = colCache.getAddress(r, c);
    var row = this._rows[address.row-1];
    return row ? row.findCell(address.col) : undefined;
  },

  // return the cell at [r,c] or address given by r. If not found, create a new one.
  getCell: function(r, c) {
    var address = colCache.getAddress(r, c);
    var row = this.getRow(address.row);
    return row._getCell(address);
  },

  // =========================================================================
  // Merge

  // convert the range defined by ['tl:br'], [tl,br] or [t,l,b,r] into a single 'merged' cell
  mergeCells: function() {
    var dimensions = new Range(Array.prototype.slice.call(arguments, 0)); // convert arguments into Array

    // check cells aren't already merged
    _.each(this._merges, function(merge) {
      if (merge.intersects(dimensions)) {
        throw new Error('Cannot merge alreay merged cells');
      }
    });

    // apply merge
    var master = this.getCell(dimensions.top, dimensions.left);
    for (var i = dimensions.top; i <= dimensions.bottom; i++) {
      for (var j = dimensions.left; j <= dimensions.right; j++) {
        // merge all but the master cell
        if ((i > dimensions.top) || (j > dimensions.left)) {
          this.getCell(i,j).merge(master);
        }
      }
    }

    // index merge
    this._merges[master.address] = dimensions;
  },
  _unMergeMaster: function(master) {
    // master is always top left of a rectangle
    var merge = this._merges[master.address];
    if (merge) {
      for (var i = merge.top; i <= merge.bottom; i++) {
        for (var j = merge.left; j <= merge.right; j++) {
          this.getCell(i,j).unmerge();
        }
      }
      delete this._merges[master.address];
    }
  },

  // scan the range defined by ['tl:br'], [tl,br] or [t,l,b,r] and if any cell is part of a merge,
  // un-merge the group. Note this function can affect multiple merges and merge-blocks are
  // atomic - either they're all merged or all un-merged.
  unMergeCells: function() {
    var dimensions = new Range(Array.prototype.slice.call(arguments, 0)); // convert arguments into Array

    // find any cells in that range and unmerge them
    for (var i = dimensions.top; i <= dimensions.bottom; i++) {
      for (var j = dimensions.left; j <= dimensions.right; j++) {
        var cell = this.findCell(i,j);
        if (cell) {
          if (cell.type === Enums.ValueType.Merge) {
            // this cell merges to another master
            this._unMergeMaster(cell.master);
          } else if (this._merges[cell.address]) {
            // this cell is a master
            this._unMergeMaster(cell);
          }
        }
      }
    }
  },
  // ===========================================================================
  // Deprecated
  get tabColor() {
    console.trace('worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor');
    return this.properties.tabColor;
  },
  set tabColor(value) {
    console.trace('worksheet.tabColor property is now deprecated. Please use worksheet.properties.tabColor');
    return this.properties.tabColor = value;
  },

  // ===========================================================================
  // Model

  get model() {
    var model = {
      id: this.id,
      name: this.name,
      dataValidations: this.dataValidations.model,
      properties:  this.properties.model
    };

    // =================================================
    // columns
    model.cols = Column.toModel(this.columns);

    // ==========================================================
    // Rows
    var rows = model.rows = [];
    var dimensions = model.dimensions = new Range();
    _.each(this._rows, function(row) {
      var rowModel = row.model;
      if (rowModel) {
        dimensions.expand(rowModel.number, rowModel.min, rowModel.number, rowModel.max);
        rows.push(rowModel);
      }
    });

    // ==========================================================
    // Merges
    model.merges = [];
    _.each(this._merges, function(merge) {
      model.merges.push(merge.range);
    });

    return model;
  },
  _parseRows: function(model) {
    var self = this;
    this._rows = [];
    _.each(model.rows, function(rowModel) {
      var row = new Row(self, rowModel.number);
      self._rows[row.number-1] = row;
      row.model = rowModel;
    });
  },
  _parseMergeCells: function(model) {
    var self = this;
    _.each(model.mergeCells, function(merge) {
      self.mergeCells(merge);
    });
  },
  set model(value) {
    this.name = value.name;
    this._columns = Column.fromModel(this, value.cols);
    this._parseRows(value);
    
    this._parseMergeCells(value);
    this.dataValidations = new DataValidations(value.dataValidations);
    this.properties.model = value.properties;
  }
};

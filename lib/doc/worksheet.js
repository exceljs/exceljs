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

var _ = require('../utils/under-dash');


var utils = require('./../utils/utils');
var colCache = require('./../utils/col-cache');
var Range = require('./range');
var Row = require('./row');
var Column = require('./column');
var Enums = require('./enums');
var DataValidations = require('./data-validations');

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
  this.properties = Object.assign({}, {
    defaultRowHeight: 15,
    dyDescent: 55,
    outlineLevelCol: 0,
    outlineLevelRow: 0
  }, options.properties);
  
  // for all things printing
  this.pageSetup = Object.assign({}, {
    margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer:  0.3 },
    orientation: 'portrait',
    horizontalDpi: 4294967295,
    verticalDpi: 4294967295,
    fitToPage: !!(options.pageSetup && ((options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) && !options.pageSetup.scale)),
    pageOrder: 'downThenOver',
    blackAndWhite: false,
    draft: false,
    cellComments: 'None',
    errors: 'displayed',
    scale: 100,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: undefined,
    showRowColHeaders: false,
    showGridLines: false,
    firstPageNumber: undefined,
    horizontalCentered: false,
    verticalCentered: false,
    rowBreaks: null,
    colBreaks: null
  }, options.pageSetup);

  this.dataValidations = new DataValidations();

  // for freezepanes, split, zoom, gridlines, etc
  this.views = options.views || [];
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
    this._rows.forEach(function(row) {
      if (row) {
        var rowDims = row.dimensions;
        if (rowDims) {
          dimensions.expand(row.number, rowDims.min, row.number, rowDims.max);
        }
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
    value.forEach(function(defn) {
      var column = new Column(self, count++, false);
      _columns.push(column);
      column.defn = defn;
    });
  },

  // get a single column by col number. If it doesn't exist, create it and any gaps before it
  getColumn: function(c) {
    if (typeof c == 'string'){
      // if it matches a key'd column, return that
      var col = this._keys[c];
      if (col) return col;

      // otherwise, assume letter
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
  spliceColumns: function(start, count) {
    // each member of inserts is a column of data. 

    var inserts = Array.prototype.slice.call(arguments, 2);
    var _rows = this._rows;
    var nRows = _rows.length;
    if (inserts.length > 0) {
      // must iterate over all rows whether they exist yet or not
      for (var i = 0; i <  nRows; i++) {
        var rowArguments = [start, count];
        inserts.forEach(function(insert) {
          rowArguments.push(insert[i] || null);
        });
        var row = this.getRow(i+1);
        row.splice.apply(row, rowArguments);
      }
    } else {
      // nothing to insert, so just splice all rows
      this._rows.forEach(function(row, index) {
        if (row) { row.splice(start, count); }
      });
    }
    
    // splice column definitions
    var nExpand = inserts.length - count;
    var nKeep = start + count;
    var nEnd = this._columns.length;
    if (nExpand < 0) {
      for (i = nKeep; i <= nEnd; i++) {
        this.getColumn(i + nExpand).defn = this.getColumn(i).defn
      }
      for (i = nEnd + nExpand + 1; i <= nEnd; i++) {
        this.getColumn(i).defn = null;
      }
    } else {
      for (i = nEnd; i >= nKeep; i--) {
        this.getColumn(i + nExpand).defn = this.getColumn(i).defn;
      }
    }
    for (i = start; i < nKeep + nExpand; i++) {
      this.getColumn(i).defn = null;
    }
  },

  get columnCount() {
    var maxCount = 0;
    this.eachRow(function(row) {
      maxCount = Math.max(maxCount, row.cellCount);
    });
    return maxCount;
  },
  get actualColumnCount() {
    // performance nightmare - for each row, counts all the columns used
    var counts = [];
    var count = 0;
    this.eachRow(function(row) {
      row.eachCell(function(cell) {
        var col = cell.col;
        if (!counts[col]) {
          counts[col] = true;
          count++;
        }
      });
    });
    return count;
  },


  // =========================================================================
  // Rows

  _commitRow: function() {
    // nop - allows streaming reader to fill a document
  },

  get _lastRowNumber() {
    // need to cope with results of splice
    var _rows = this._rows;
    var n = _rows.length;
    while ((n > 0) && (_rows[n-1] === undefined)) {
      n--;
    }
    return n;
  },
  get _nextRow() {
    return this._lastRowNumber + 1;
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

  get rowCount() {
    return this._lastRowNumber;
  },
  get actualRowCount() {
    // counts actual rows that have actual data
    var count = 0;
    this.eachRow(function() {
      count++;
    });
    return count;
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
  
  spliceRows: function(start, count) {
    // same problem as row.splice, except worse.
    var inserts = Array.prototype.slice.call(arguments, 2);
    var nKeep = start + count;
    var nExpand = inserts.length - count;
    var nEnd = this._rows.length;
    var i, rSrc, rDst;
    if (nExpand < 0) {
      // remove rows
      for (i = nKeep; i <= nEnd; i++) {
        rSrc = this._rows[i-1];
        if (rSrc) {
          this.getRow(i+nExpand).values = rSrc.values;
          this._rows[i-1] = undefined;
        } else {
          this._rows[i+nExpand-1] = undefined;
        }
      }
    } else if (nExpand > 0) {
      // insert new cells
      for (i = nEnd; i >= nKeep; i--) {
        rSrc = this._rows[i-1];
        if (rSrc) {
          this.getRow(i+nExpand).values = rSrc.values;
        } else {
          this._rows[i+nExpand-1] = undefined;
        }
      }
    }

    // now copy over the new values
    for (i = 0; i < inserts.length; i++) {
      this.getRow(start + i).values = inserts[i];
    }
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
        if (row && row.hasValues) {
          iteratee(row, row.number);
        }
      });
    }
  },

  // return all rows as sparse array
  getSheetValues: function() {
    var rows = [];
    this._rows.forEach(function(row) {
      if (row) { rows[row.number] = row.values; }
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
  
  get hasMerges() {
    return _.some(this._merges, function(merge, address) {
      return true;
    });
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
      properties: this.properties,
      pageSetup: this.pageSetup,
      views: this.views
    };

    // =================================================
    // columns
    model.cols = Column.toModel(this.columns);

    // ==========================================================
    // Rows
    var rows = model.rows = [];
    var dimensions = model.dimensions = new Range();
    this._rows.forEach(function(row) {
      var rowModel = row && row.model;
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
    model.rows.forEach(function(rowModel) {
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
    this.properties = value.properties;
    this.pageSetup = value.pageSetup;
    this.views = value.views;
  }
};

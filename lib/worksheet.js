/**
 * Copyright (c) 2014 Guyon Roche
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:</p>
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
"use strict";

var _ = require("underscore");
var Promise = require("bluebird");

var utils = require("./utils");
var colCache = require("./colcache");
var Dimensions = require("./dimensions");
var Row = require("./row");
var Column = require("./column");
var Cell = require("./cell");

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
    this.name = options.name || "Sheet" + this.id;
    
    // cells is a "sparse array" of Cell objects indexed by address (e.g. A1)
    this._cells = {};
    
    // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
    // Note: _rows is zero based. Must subtract 1 to go from cell.row to index
    this._rows = [];
    
    // column definitions    
    this._columns = null;
    
    // column keys (addRow convenience)
    this._keys = {};
    
    // keep record of all merges
    this._merges = {};
    
    this._workbook = options.workbook;
}

Worksheet.prototype = {
    
    // when you're done with this worksheet, call this to remove from workbook
    destroy: function() {
        this._workbook._removeWorksheet(this);
    },
    
    get dimensions() {
        var dimensions = new Dimensions();
        _.each(this._rows, function(row, index) {
            var rowDims = row.dimensions;
            if (rowDims) {
                dimensions.expand(row.number, rowDims.min, row.number, rowDims.max);
            }
        });
        return dimensions;
    },

    
    // set column properties
    // array of column definitions including: header(s), width, style, key
    // if used, should be set before any rows or cells are assigned, otherwise, row data may be overwritten or left stale
    get columns() {
        return this._columns;
    },
    set columns(value) {
        var self = this;
        this._columns = value.map(function(defn) { return new Column(defn); });
        
        // post-process. Look for headers, justify, keys, etc.
        this._headerRowCount = this._columns
            .reduce(function(pv,cv){ return Math.max(pv, cv.headerCount)}, 0);
            
        // process header rows
        if (this._headerRowCount) {
            var justifyTop = true;
            _.each(this._columns, function(column, index) {
                var j = index + 1;
                var count = column.headerCount;
                var start = justifyTop ? 1 : this._headerRowCount + 1 - count;
                for (var i = start; i < start + count; i++) {
                    var cell = self.getCell(i, j);
                    cell.value = column.headers[i-start];
                }
            });
        }
        
        // assign column keys
        _.each(this._columns, function(column, index) {
            if (column.key) {
                self._keys[column.key] = index;
            }
        });
    },
    get _nextRow() {
        return this._rows.length + 1;
    },
    getRow: function(r) {
        var row = this._rows[r-1];
        if (!row) {
            row = this._rows[r-1] = new Row({number: r});
        }
        return row;
    },
    findCell: function(r, c) {
        var address = r;
        if (c) {
            address = colCache.n2l(c) + r;
        }
        return this._cells[address];
    },
    getCell: function(r, c) {
        var address = r;
        if (c) {
            address = colCache.n2l(c) + r;
        }
        var cell = this._cells[address];
        if (cell) { return cell }
        cell = this._cells[address] = new Cell({
            address: address,
            worksheet: this
        });
        
        var row = this.getRow(cell.row);
        row.setCell(cell);
        
        return cell;
    },
    
    addRow: function(value) {
        var self = this;
        var row = this._nextRow;
        if (value instanceof Array) {
            _.each(value, function(item, index) {
                self.getCell(row, index + 1).value = item;
            });
        } else {
            // assume object with column keys
            _.each(this._keys, function(index, key) {
                if (value[key] !== undefined) {
                    self.getCell(row, index + 1).value = value[key];
                }
            });
        }
    },
    mergeCells: function() {
        var dimensions = new Dimensions(Array.prototype.slice.call(arguments, 0)); // convert arguments into Array
        
        // check cells aren't already merged
        for (var i = dimensions.top; i <= dimensions.bottom; i++) {
            for (var j = dimensions.left; j <= dimensions.right; j++) {
                var cell = this.findCell(i,j);
                if (cell && cell.isMerged) {
                    throw new Error("Cannot merge alreay merged cells");
                }
            }
        }
        
        // apply merge
        var master = this.getCell(dimensions.top, dimensions.left);
        for (var i = dimensions.top; i <= dimensions.bottom; i++) {
            for (var j = dimensions.left; j <= dimensions.right; j++) {
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
    unMergeCells: function() {
        var dimensions = this._getDimensions(arguments);
        
        // find any cells in that range and unmerge them
        for (var i = dimensions.top; i <= dimensions.bottom; i++) {
            for (var j = dimensions.left; j <= dimensions.right; j++) {
                var cell = this.findCell(i,j);
                if (cell) {
                    if (cell.isMergedTo()) {
                        this._unMergeMaster(cell.master);
                    } else {
                        this._unMergeMaster(cell);
                    }
                }
            }
        }
    },
    
    // ===========================================================================
    // Model
    
    get model() {
        var self = this;
        var model = {
            id: this.id,
            name: this.name
        };
        
        // =================================================
        // columns
        var cols = [];
        var col = null;
        _.each(self.columns, function(column, index) {
            if (!col || !column.equivalentTo(col)) {
                col = {
                    min: index + 1,
                    max: index + 1,
                    width: column.width
                };
                // only add to list if col defn not default
                if (!column.isDefault) {
                    cols.push(col);
                }
            } else {
                col.max = index + 1;
            }
        });
        if (cols.length) {
            model.cols = cols;
        }
        
        // ==========================================================
        // Rows
        var rows = model.rows = [];
        var dimensions = model.dimensions = new Dimensions();
        _.each(this._rows, function(row, index) {
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
    
    _parseCols: function(model) {
        if (model.cols) {
            this._columns = [];
            var count = 1;
            var index = 0;
            while (index < model.cols.length) {
                var col = model.cols[index++];
                while (count < col.min) {
                    this._columns.push(new Column());
                    count++;
                }
                while (count <= col.max) {
                    this._columns.push(new Column(col));
                    count++;
                }
            }
        } else {
            this._columns = null;
        }
    },
    _parseRows: function(model) {
        var self = this;
        this._rows = [];
        this._cells = {};
        _.each(model.rows, function(rowModel) {
            var row = new Row(rowModel);
            self._rows[row.number-1] = row;
            _.each(rowModel.cells, function(cellModel) {
                switch(cellModel.type) {
                    case Cell.Types.Null:
                    case Cell.Types.Merge:
                        // special case - don't add these types
                        break;
                    default:
                        self.getCell(cellModel.address).model = cellModel;
                        break;
                }
            });
        });
    },
    _parseMergeCells: function(model) {
        var self = this;
        _.each(model.merges, function(merge) {
            self.mergeCells(merge);
        });
    },
    set model(value) {
        this.name = value.name;
        this._parseCols(value);
        this._parseRows(value);
        this._parseMergeCells(value);
    }
}

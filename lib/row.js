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
"use strict";

var _ = require("underscore");
var Enums = require("./enums");
var colCache = require("./colcache");
var Cell = require("./cell");

var Row = module.exports = function(worksheet, number) {
    this._worksheet = worksheet;    
    this._number = number;
    this._cells = [];
}

Row.prototype = {
    // return the row number
    get number() {
        return this._number;
    },
    
    // given {address, row, col}, find or create new cell
    _getCell: function(address) {
        var cell = this._cells[address.col-1];
        if (!cell) {
            cell = new Cell(this, address.address);
            this._cells[address.col-1] = cell;
            this._worksheet._cells[address.address] = cell;
        }
        return cell;
    },
    
    // get cell by key, letter or column number
    getCell: function(col) {
        if (typeof col == "string") {
            // is it a key?
            var column = this._worksheet._keys[col];
            if (column) {
                col = column.number;
            } else {
                col = colCache.l2n(col);
            }
        }
        return this._cells[col-1] ||
                this._getCell({
                    address: colCache.encodeAddress(this._number, col),
                    row: this._number,
                    col: col
                });
    },
    
    // Iterate over all non-null cells in this row
    eachCell: function(iteratee) {
        _.each(this._cells, function(cell, index) {
            if (cell.type != Enums.ValueType.Null) {
                iteratee(cell, index + 1);
            }
        });
    },
    
    // return a sparse array of cell values
    get values() {
        var values = [];
        _.each(this._cells, function(cell) {
            if (cell.type != Enums.ValueType.Null) {
                values[cell.col] = cell.value;
            }
        });
        return values;
    },
    
    // set the values by contiguous or sparse array, or by key'd object literal
    set values(value) {
        var self = this;
        
        // this operation is not additive - any prior cells are removed
        _.each(this._cells, function(cell) {
            delete self._worksheet._cells[cell.address];
        });
        this._cells = [];
        
        if (value instanceof Array) {
            var offset = 0;
            if (value[0] != undefined) {
                // contiguous array - start at column 1
                offset = 1;
            }
            _.each(value, function(item, index) {
                self._getCell({
                    address: colCache.encodeAddress(self._number, index+offset),
                    row: self._number,
                    col: index+offset
                }).value = item;
            });
        } else {
            // assume object with column keys
            _.each(this._worksheet._keys, function(column, key) {
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
        return _.some(this._cells, function(cell) {
            return cell.type != Enums.ValueType.Null;
        });
    },
    
    // get the min and max column number for the non-null cells in this row or null
    get dimensions() {
        var min = 0;
        var max = 0;
        _.each(this._cells, function(cell) {
            if (cell.type != Enums.ValueType.Null) {
                if (!min || (min > cell.col)) { min = cell.col; }
                if (max < cell.col) { max = cell.col; }
            }
        });
        return min > 0 ?{
            min: min,
            max: max
        } : null;
    },
    
    get model() {
        var cells = [];
        var min = 0;
        var max = 0;
        _.each(this._cells, function(cell) {
            if (cell.type != Enums.ValueType.Null) {
                var cellModel = cell.model;
                if (cellModel) {
                    if (!min || (min > cell.col)) { min = cell.col; }
                    if (max < cell.col) { max = cell.col; }
                    cells.push(cellModel);
                }
            }
        });
        
        return (this.height || cells.length) ? {
            cells: cells,
            number: this.number,
            min: min,
            max: max,
            height: this.height
        }: null;
    },
    set model(value) {
        if (value.number != this._number) {
            throw new Error("Invalid row number in model");
        }
        var self = this;
        this._cells = [];
        _.each(value.cells, function(cellModel) {
            //console.log(JSON.stringify(colCache.decodeAddress(cellModel.address)) + " ==> " + JSON.stringify(cellModel));
            switch(cellModel.type) {
                case Cell.Types.Null:
                case Cell.Types.Merge:
                    // special case - don't add these types
                    break;
                default:
                    self._getCell(colCache.decodeAddress(cellModel.address)).model = cellModel;
                    break;
            }
        });
        
        if (value.height) {
            this.height = value.height;
        } else {
            delete this.height;
        }
    }
}

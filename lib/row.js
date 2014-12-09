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
var Enums = require("./enums");

// Cell requirements
//  Operate inside a worksheet
//  Store and retrieve a value with a range of types: text, number, date, hyperlink, reference, formula, etc.
//  Manage/use and manipulate cell format either as local to cell or inherited from column or row.

var Row = module.exports = function(options) {
    this._number = options.number;
    this._cells = [];
}

Row.prototype = {
    // return the row number
    get number() {
        return this._number;
    },
    
    setCell: function(cell) {
        this._cells[cell.col-1] = cell;
    },
    getCell: function(col) {
        return this._cells[col-1];
    },
    
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
            var cellModel = cell.model;
            if (cellModel) {
                if (!min || (min > cell.col)) { min = cell.col; }
                if (max < cell.col) { max = cell.col; }
                cells.push(cellModel);
            }
        });
        
        return cells.length ? {
            cells: cells,
            number: this.number,
            min: min,
            max: max
        }: null;
    }
}

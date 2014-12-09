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
var utils = require("./utils");
var Worksheet = require("./worksheet");

var XLSX = require("./xlsx/xlsx");

// Workbook requirements
//  Load and Save from file and stream
//  Access/Add/Delete individual worksheets
//  Manage String table, Hyperlink table, etc.
//  Manage scaffolding for contained objects to write to/read from

var Workbook = module.exports = function() {
    this._worksheets = [];
}

Workbook.prototype = {
    get xlsx() {
        return this._xlsx || (this._xlsx = new XLSX(this));
    },
    get nextId() {
        // find the next unique spot to add worksheet
        var i;
        for (i = 1; i < this._worksheets.length; i++) {
            if (!this._worksheets[i]) {
                return i;
            }
        }
        return this._worksheets.length || 1;
    },
    addWorksheet: function(name) {
        
        var id = this.nextId;
        name = name || "sheet" + id;
        
        var worksheet = new Worksheet({
            id: id,
            name: name,
            workbook: this
        });
        
        this._worksheets.push(worksheet);
        return worksheet;
    },
    _removeWorksheet: function(worksheet) {
        delete this._worksheets[worksheet.id];
    },
    removeWorksheet: function(id) {
        var worksheet = this.getWorksheet(id);
        if (worksheet) {
            worksheet.destroy();
        }
    },
    
    getWorksheet: function(id) {
        if (typeof(id) === "number") {
            return this._worksheets[id];
        } else {
            return _.find(this._worksheets, function(worksheet) {
                return worksheet.name == id;
            });
        }
    },
    get worksheets() {
        // return a clone of _worksheets
        return this._worksheets.filter(function(worksheet) { return worksheet; });
    },
    
    get model() {
        return {
            worksheets: this._worksheets.map(function(worksheet) { return worksheet.model; })
        };
    },
    set model(value) {
        var self = this;
        
        this._worksheets = [];
        _.each(value.worksheets, function(worksheetModel) {
            var id = worksheetModel.id;
            var name = worksheetModel.name;            
            var worksheet = self._worksheets[id] = new Worksheet({
                id: id,
                name: name,
                workbook: self
            });
            
            worksheet.model = worksheetModel; 
        });
    }
}
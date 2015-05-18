/**
 * Copyright (c) 2015 Guyon Roche
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

var fs = require("fs");
var Promise = require("bluebird");
var _ = require("underscore");

var utils = require("../../utils/utils");
var colCache = require("../../utils/col-cache");
var Enums = require("../../enums");
var Dimensions = require("../../utils/dimensions");

var Row = require("../../row");
var Column = require("../../column");

var StyleManager = require("../../xlsx/stylemanager");
var HyperlinkWriter = require("./hyperlink-writer");

var WorksheetWriter = module.exports = function(options) {
    // in a workbook, each sheet will have a number
    this.id = options.id;
    
    // and a name
    this.name = options.name || "Sheet" + this.id;
    
    // rows are stored here while they need to be worked on.
    // when they are comitted, they will be deleted.
    this._rows = [];
    
    // column definitions    
    this._columns = null;
    
    // column keys (addRow convenience): key ==> this._collumns index
    this._keys = {};
    
    // keep record of all merges
    this._merges = [];
    
    // keep record of all hyperlinks
    this._hyperlinkWriter = new HyperlinkWriter(options);
    
    // keep a record of dimensions
    this._dimensions = new Dimensions();
    
    // count of added rows
    this._committedRow = 0;
    this._nextRow = 1;
    
    // committed flag
    this.committed = false;
    
    // using shared strings creates a smaller xlsx file but may use more memory
    this.useSharedStrings = options.useSharedStrings || false;

    this._workbook = options.workbook;
    
    // start writing to stream now
    this._writeOpenWorksheet();
    
    this.startedData = false;
}
WorksheetWriter.prototype = {
    get stream() {
        if (!this._stream) {
            this._stream = this._workbook._openStream('/xl/worksheets/sheet' + this.id + '.xml');
            //this._stream = fs.createWriteStream("./out/sheet1.xml")
        }
        return this._stream;
    },

    // destroy - not a valid operation for a streaming writer
    // even though some streamers might be able to, it's a bad idea.
    destroy: function() {
        throw new Error("Invalid Operation: destroy");
    },
    
    commit: function() {
        if (this.committed) {
            return;
        }
        var self = this;
        // commit all rows
        _.each(this._rows, function(crow) {
            // write the row to the stream
            self._writeRow(crow);
        });
        
        // we _cannot_ accept new rows from now on
        this._rows = null;
        
        this._writeCloseSheetData();
        this._writeMergeCells();
        
        // for some reason, Excel can't handle dimensions at the bottom of the file
        //this._writeDimensions();
        
        this._writeHyperlinks();
        this._writePageMargins();
        this._writeCloseWorksheet();
        
        // signal end of stream to workbook
        this.stream.end();
        
        // also commit the hyperlinks if any
        this._hyperlinkWriter.commit();
        
        this.committed = true;
    },
    
    // return the current dimensions of the writer
    get dimensions() {
        return this._dimensions;
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
        if (typeof c == "string"){
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
    _commitRow: function(row) {
        var self = this;
        if (this._rows[row.number-1] === row) {
            var rows = [];
            _.every(this._rows, function(crow) {
                // remember this row 
                rows.push(crow);
                
                // every will short circuit once we've reached row (and return false)
                return crow !== row;
            });
            _.each(rows, function(crow) {
                // write the row to the stream
                self._writeRow(crow);
            });
        }
    },
    
    get lastRow() {
        // returns last uncommitted row
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
    
    getRow: function(rowNumber) {
        // may fail if rows have been comitted
        if (rowNumber <= this._committedRow) {
            throw new Error("Out of bounds: this row has been committed");
        }
        var row = this._rows[rowNumber-1];
        if (!row) {
            this._rows[rowNumber-1] = row = new Row(this, rowNumber);
            this._nextRow = rowNumber + 1;
        }
        return row;
    },
    
    addRow: function(value) {
        var rowNumber = this._nextRow++;
        var row = new Row(this, rowNumber);
        this._rows[rowNumber-1] = row;
        row.values = value;
        return row;
    },
    
    // ================================================================================
    // Cells
    
    // returns the cell at [r,c] or address given by r. If not found, return undefined
    findCell: function(r, c) {
        var address = colCache.getAddress(r, c);
        var row = this._rows[address.row-1];
        return row ? row.findCell(address.column) : undefined;
    },
    
    // return the cell at [r,c] or address given by r. If not found, create a new one.
    getCell: function(r, c) {
        var address = colCache.getAddress(r, c);
        var row = this.getRow(address.row);
        return row._getCell(address);
    },    
    
    mergeCells: function() {
        // may fail if rows have been comitted
        var dimensions = new Dimensions(Array.prototype.slice.call(arguments, 0)); // convert arguments into Array
        
        // check cells aren't already merged
        _.each(this._merges, function(merge) {
            if (merge.intersects(dimensions)) {
                throw new Error("Cannot merge alreay merged cells");
            }
        });
        
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
        this._merges.push(dimensions);
    },
    
    // ================================================================================
    _writeOpenWorksheet: function() {
        this.stream.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        this.stream.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">');
        
        this.stream.write('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
        this.stream.write('<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>');
    },
    _writeColumns: function() {
        var self = this;
        var cols = Column.toModel(this.columns);
        if (cols) {
            var xml = [];
            xml.push('<cols>');
            _.each(cols, function(col) {
                xml.push('<col');
                xml.push(' min="' + col.min + '"');
                xml.push(' max="' + col.max + '"');
                xml.push(' width="' + col.width + '"');
                if (col.style) {
                    var colStyleId = self._workbook.styles.addStyleModel(col.style);
                    if (colStyleId) {
                        xml.push(' style="' + colStyleId + '"');
                    }
                }
                if (col.isCustomWidth) {
                    xml.push(' customWidth="1"');
                }
                xml.push('/>');
            });
            xml.push('</cols>');
            this.stream.write(xml.join(''));
        }
    },
    _writeOpenSheetData: function() {
        this.stream.write('<sheetData>');
    },
    _writeRow: function(row) {
        var self = this;
        
        if (!this.startedData) {
            this._writeColumns();
            this._writeOpenSheetData();
            this.startedData = true;
        }
        
        if (row.hasValues || row.height) {
            var rowStyleId = this._workbook.styles.addStyleModel(row.style);
            var rowDimensions = row.dimensions;
            this._dimensions.expand(row.number, rowDimensions.min, row.number, rowDimensions.max);
            var xml = [];
            xml.push('<row r="' + row.number + '"');
            if (row.height) {
                xml.push(' ht="' + row.height + '" customHeight="1"');
            }
            xml.push(' spans="' + rowDimensions.min + ':' + rowDimensions.max + '"');
            if (row.styleId) {
                xml.push(' s="' + row.styleId + ' customFormat="1"');
            }
            xml.push(' x14ac:dyDescent="0.25">');
            
            row.eachCell(function(cell) {
                var xmlCell = ['<c r="' + cell.address + '"'];
                var cellStyleId = self._workbook.styles.addStyleModel(cell.style, cell.effectiveType);
                if (cellStyleId) {
                    xmlCell.push(' s="' + cellStyleId + '"');
                }
                switch(cell.type) {
                    case Enums.ValueType.Null:
                        xmlCell = [];
                        break;
                    case Enums.ValueType.Merge:
                        if (cellStyleId) { xmlCell.pop(); }
                        xmlCell.push('/>');
                        break;
                    case Enums.ValueType.Number:
                        xmlCell.push('><v>' + cell.value + '</v></c>');
                        break;
                    case Enums.ValueType.String:
                        if (self.useSharedStrings) {
                            var stringId = self._workbook.sharedStrings.add(cell.value);
                            xmlCell.push(' t="s"><v>' + stringId + '</v></c>');
                        } else {
                            xmlCell.push(' t="str"><v>' + utils.xmlEncode(cell.value) + '</v></c>');
                        }
                        break;
                    case Enums.ValueType.Date:
                        xmlCell.push('><v>' + utils.dateToExcel(cell.value) + '</v></c>');
                        break;
                    case Enums.ValueType.Hyperlink:
                        if (self.useSharedStrings) {
                            var stringId = self._workbook.sharedStrings.add(cell.value.text);
                            xmlCell.push(' t="s"><v>' + stringId + '</v></c>');
                        } else {
                            xmlCell.push(' t="str"><v>' + utils.xmlEncode(cell.value.text) + '</v></c>');
                        }
                        self._hyperlinkWriter.add({
                            address: cell.address,
                            target: utils.xmlEncode(cell.value.hyperlink)
                        });
                        break;
                    case Enums.ValueType.Formula:
                        switch(cell.effectiveType) {
                            case Enums.ValueType.Null:
                                xmlCell.push('><f>' + utils.xmlEncode(cell.value.formula) + '</f></c>');
                                break;
                            case Enums.ValueType.String:
                                // oddly, formula results don't ever use shared strings
                                xmlCell.push(' t="str"><f>' + utils.xmlEncode(cell.value.formula) + '</f>');
                                xmlCell.push('<v>' + utils.xmlEncode(cell.value.result) + '</v>');
                                xmlCell.push('</c>');
                                break;
                            case Enums.ValueType.Number:
                                xmlCell.push('><f>' + utils.xmlEncode(cell.value.formula) + '</f>');
                                xmlCell.push('<v>' + cell.value.result + '</v>');
                                xmlCell.push('</c>');
                                break;
                            case Enums.ValueType.Date:
                                xmlCell.push('><f>' + utils.xmlEncode(cell.value.formula) + '</f>');
                                xmlCell.push('<v>' + utils.dateToExcel(cell.value.result) + '</v>');
                                xmlCell.push('</c>');
                                break;
                            case Enums.ValueType.Hyperlink: // ??
                            case Enums.ValueType.Formula:
                            default:
                                throw new Error("I could not understand type of value");
                        }
                        break;
                }
                xml.push(xmlCell.join(''));
            });
            xml.push('</row>');
            this.stream.write(xml.join(''));
        }
        
        delete this._rows[row.number-1];
        this._committedRow = Math.max(this._committedRow, row.number);
    },
    _writeCloseSheetData: function() {
        this.stream.write('</sheetData>');
    },
    _writeMergeCells: function() {
        var self = this;
        
        if (this._merges.length) {
            this.stream.write('<mergeCells count="' + this._merges.length + '">');
            _.each(this._merges, function(merge) {
                self.stream.write('<mergeCell ref="' + merge + '"/>');
            });
            this.stream.write('</mergeCells>');
        }
    },
    _writeHyperlinks: function() {
        var self = this;
        if (this._hyperlinkWriter.count) {
            this.stream.write('<hyperlinks>');
            this._hyperlinkWriter.each(function(hyperlink) {
                self.stream.write('<hyperlink ref="' + hyperlink.address + '" r:id="' + hyperlink.rId + '"/>');
            });
            this.stream.write('</hyperlinks>');
        }
    },
    _writePageMargins: function() {
        this.stream.write('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
    },
    _writeDimensions: function() {
        this.stream.write('<dimension ref="' + this._dimensions + '"/>');
    },
    _writeCloseWorksheet: function() {
        this.stream.write('</worksheet>');
    }
}
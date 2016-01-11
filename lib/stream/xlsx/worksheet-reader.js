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
var events = require("events");
var Promise = require("bluebird");
var _ = require("underscore");
var Sax = require("sax");

var utils = require("../../utils/utils");
var colCache = require("../../utils/col-cache");
var Enums = require("../../enums");
var Dimensions = require("../../utils/dimensions");

var Row = require("../../row");
var Column = require("../../column");

var WorksheetReader = module.exports = function(workbook, id) {
    this.workbook = workbook;
    this.id = id;
    
    // and a name
    this.name = "Sheet" + this.id;
    
    this.rowCount = 0;

    // column definitions    
    this._columns = null;
    this._keys = {};
    
    // keep a record of dimensions
    this._dimensions = new Dimensions();
}

utils.inherits(WorksheetReader, events.EventEmitter, {

    // destroy - not a valid operation for a streaming writer
    // even though some streamers might be able to, it's a bad idea.
    destroy: function() {
        throw new Error("Invalid Operation: destroy");
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
    // Read
    
    _emitRow: function(row) {
        if (this.listenerCount("row")) this.emit("row", row);
    },
    
    read: function(entry, hyperlinkReader) {
        var self = this;
        
        // var emitSheet = false;
        // var emitHyperlinks = false;
        // var hyperlinks = null;

        // if (options.worksheets.indexOf("emit") > -1) {
        //     emitSheet = true;
        // }

        // if (options.hyperlinks.indexOf("emit") > -1) {
        //     emitHyperlinks = true;
        // }
        // if (options.hyperlinks.indexOf("cache") > -1) {
        //     this.hyperlinks = hyperlinks = {};
        // }

        // if (!self.listenerCount("row") && !self.listenerCount("rowCount") && !self.listenerCount("hyperlink")) {
        //     entry.autodrain();
        //     return;
        // }
        
        // references
        //var sharedStrings = this.workbook.sharedStrings;
        var styles = this.workbook.styles;
        
        // xml position
        var inCols = false;
        var inRows = false;
        var inHyperlinks = false;
        
        // parse state
        var cols = null;
        var row = null;
        var c = null;
        var current = null;
        
        var parser = Sax.createStream(true, {});
        parser.on('opentag', function(node) {
            //if (self.listenerCount("row") || self.listenerCount("colCount") || self.listenerCount("rowCount")) {
            switch (node.name) {
                case "cols":
                    inCols = true;
                    cols = [];
                    break;
                case "sheetData":
                    inRows = true;
                    break;
                
                case "col":
                    if (inCols) {
                        cols.push({
                            min: parseInt(node.attributes.min),
                            max: parseInt(node.attributes.max),
                            width: parseFloat(node.attributes.width),
                            styleId: parseInt(node.attributes.style || "0")
                        });
                    }
                    break;
                
                case "row":
	        		++self.rowCount
                    if (inRows && self.listenerCount("row")) {
                        var r = parseInt(node.attributes.r);
                        row = new Row(self, r);
                        if (node.attributes.ht) {
                            row.height = parseFloat(node.attributes.ht);
                        }
                        if (node.attributes.s) {
                            var styleId = parseInt(node.attributes.s);
                            row.style = styles.getStyleModel(styleId);
                        }
                    }
                    break;
                case "c":
                    if (row) {
                        c = {
                            ref: node.attributes.r,
                            s: parseInt(node.attributes.s),
                            t: node.attributes.t
                        };
                    }
                    break;
                case "f":
                    if (c) {
                        current = c.f = { text: "" };
                    }
                    break;
                case "v":
                    if (c) {
                        current = c.v = { text: "" };
                    }
                    break;
                case "mergeCell":
            }
            //}
            
            
            // =================================================================
            // 
            if (self.listenerCount("hyperlink")) {
                switch (node.name) {
                    case "hyperlinks":
                        inHyperlinks = true;
                        break;
                    case "hyperlink":
                        if (inHyperlinks) {
                            var hyperlink = {
                                ref: node.attributes.ref,
                                rId: node.attributes["r:id"]
                            };
                            if (emitHyperlinks) {
                                if (self.listenerCount("hyperlink")) self.emit("hyperlink", hyperlink);
                            } else {
                                hyperlinks[hyperlink.ref] = hyperlink;
                            }
                        }
                        break;
                }
            }
        });
        
        // only text data is for sheet values
        parser.on('text', function (text) {
        	if (self.listenerCount("row")) {
                if (current) {
                    current.text += text;
                }
            }
        });
        
        parser.on('closetag', function(name) {
            if (self.listenerCount("row")) {
                switch (name) {
                    case "cols":
                        inCols = false;
                        self._columns = Column.fromModel(cols);
                        break;
                    case "sheetData":
                        inRows = false;
                        break;
                    
                    case "row":
                        self._dimensions.expandRow(row);
                        self._emitRow(row);
                        row = null;
                        break;
                    
                    case "c":
                        if (row && c) {
                            var address = colCache.decodeAddress(c.ref);
                            var cell = row.getCell(address.col);
                            if (c.s) {
                                cell.style = self.workbook.styles.getStyleModel(c.s);
                            }
                            
                            if (c.f) {
                                var value = {
                                    formula: c.f.text
                                };
                                if (c.v) {
                                    if (c.t == "str") {
                                        value.result = utils.xmlDecode(c.v.text);
                                    } else {
                                        value.result = parseFloat(c.v.text);
                                    }
                                }
                                cell.value = value;
                            } else if (c.v) {
                                switch(c.t) {
                                    case "s":
                                        //console.log(c)
                                        var index = parseInt(c.v.text);
                                        cell.value = self.workbook._getSharedString(index)
                                        break;
                                    case "str":
                                        cell.value = utils.xmlDecode(c.v.text);
                                        break;
                                    default:
                                        if (utils.isDateFmt(cell.numFmt)) {
                                            cell.value = utils.excelToDate(parseFloat(c.v.text));
                                        } else {
                                            cell.value = parseFloat(c.v.text);
                                        }
                                        break;
                                }
                            }
                            c = null;
                        }
                        break;
                }
            }
            if (self.listenerCount("hyperlink")) {
                switch (name) {
                    case "hyperlinks":
                        inHyperlinks = false;
                        break;
                }
            }
        });
        parser.on('error', function (error) {
            self.emit('error', error);
        });
        parser.on('end', function() {
            self.emit('finished', self.rowCount);
        });
        
        entry.pipe(parser);
        
        // never try to control the flow
        // // create a down-stream flow-control to regulate the stream
        // var flowControl = this.workbook.flowControl.createChild();
        // flowControl.pipe(parser, {sync: true});
        // entry.pipe(flowControl);
    }
});

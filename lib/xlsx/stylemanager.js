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

var _ = require("underscore");
var Promise = require("bluebird");
var Sax = require("sax");
var Style = require("./style");
var NumFmt = require("./numfmt");
var Font = require("./font");
var Alignment = require("./alignment");

var utils = require("../utils");
var Enums = require("../enums");

// StyleManager is used to generate and parse the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
var StyleManager = module.exports = function() {
    this.styles = [];
    this.stylesIndex = {};
    
    this.models = []; // styleId -> model
    
    this.numFmts = [];
    this.numFmtHash = {}; //numFmtId -> numFmt
    this.numFmtNextId = 164; // start custom format ids here
    
    this.fonts = []; // array of font xml
    this.fontIndex = {}; // hash of xml->fontId
    
    // ---------------------------------------------------------------
    // Defaults
    
    // default (zero) font
    this._addFont(new Font({ size: 11, color: {theme:1}, name: "Calibri", family:2, scheme:"minor"}));
    
    // add default (all zero) style
    this._addStyle(new Style());
}

StyleManager.prototype = {
    // =========================================================================
    // Public Interface
    
    // addToZip - using styles.xml template, add the styles.xml file to the zip
    addToZip: function(zip) {
        var self = this;
        return utils.fetchTemplate("./xlsx/styles.xml")
            .then(function(template){
                return template(self);
            })
            .then(function(data) {
                zip.append(data, { name: "/xl/styles.xml" });
                return zip;
            });
    },
    
    // parse the styles.xml file pulling out fonts, numFmts, etc
    parse: function(stream) {
        var self = this;
        // build a style manager object from the styles.xls file
        //<cellXfs count="15">
        //    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        //    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        //</cellXfs>
        //<numFmts count="1">
        //    <numFmt numFmtId="165" formatCode="[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"/>
        //</numFmts>
        var deferred = Promise.defer();
        var parser = Sax.createStream(true, {});
        
        var inNumFmts = false;
        var inCellXfs = false;
        var inFonts = false;
        
        var style = null;
        var font = null;
        
        parser.on('opentag', function(node) {
            if (inNumFmts) {
                if (node.name == "numFmt") {
                    var numFmt = new NumFmt();
                    numFmt.parse(node);
                    self._addNumFmt(numFmt);
                }
            } else if (inCellXfs) {
                if (style) {
                    style.parse(node);
                } else if (node.name == "xf") {
                    style = new Style();
                    style.parse(node);
                }
            } else if (inFonts) {
                if (font) {
                    font.parse(node);
                } else if (node.name == "font") {
                    font = new Font();
                }
            } else {
                switch(node.name) {
                    case 'cellXfs':
                        inCellXfs = true;
                        break;
                    case 'numFmts':
                        inNumFmts = true;
                        break;
                    case 'fonts':
                        inFonts = true;
                        break;
                }
            }
        });
        parser.on('closetag', function (name) {
            if (inCellXfs) {
                switch(name) {
                    case 'cellXfs':
                        inCellXfs = false;
                        break;
                    case 'xf':
                        self._addStyle(style);
                        style = null;
                        break;
                }
            } else if (inNumFmts) {
                switch(name) {
                    case 'numFmts':
                        inNumFmts = false;
                        break;
                }
            } else if (inFonts) {
                switch(name) {
                    case 'fonts':
                        inFonts = false;
                        break;
                    case 'font':
                        self._addFont(font);
                        font = null;
                        break;
                }
            }
        });
        parser.on('end', function() {
            // warning: if style, font, inFonts, inCellXfs, inNumFmts are true!
            deferred.resolve(self);
        });
        parser.on('error', function (error) {
            deferred.reject(error);
        });
        stream.pipe(parser);
        
        return deferred.promise;
    },
    
    // add a cell's style model to the collection
    // each style property is processed and cross-referenced, etc.
    // the styleId is returned. Note: cellType is used when numFmt not defined
    addStyleModel: function(model, cellType) {
        // Note: a WeakMap --> styleId would be very useful here!
        
        var style = new Style();
        
        if (model.numFmt) {
            style.numFmtId = this._addNumFmtStr(model.numFmt);
        } else {
            switch(cellType) {
                case Enums.ValueType.Number:
                    style.numFmtId = this._addNumFmtStr("General");
                    break;
                case Enums.ValueType.Date:
                    style.numFmtId = this._addNumFmtStr("mm-dd-yy");
                    break;
            }
        }
        
        if (model.font) {
            style.fontId = this._addFont(new Font(model.font));
        }
        
        if (model.alignment) {
            style.alignment = new Alignment(model.alignment);
        }
        
        return this._addStyle(style);
    },
    
    // given a styleId (i.e. s="n"), get the cell's style model
    // objects are shared where possible.
    getStyleModel: function(id) {
        // have we built this model before?
        var model = this.models[id];
        if (model) return model;
        
        // if the style doesn't exist return null
        var style = this.styles[id];
        if (!style) return null;
        
        // build a new model
        var model = this.models[id] = {};
        
        // -------------------------------------------------------
        // number format
        var numFmt = (this.numFmtHash[style.numFmtId] && this.numFmtHash[style.numFmtId].formatCode) ||
                            NumFmt.getDefaultFmtCode(style.numFmtId);
        if (numFmt) {
            model.numFmt = numFmt;
        }
        
        // -------------------------------------------------------
        // font
        var font = this.fonts[style.fontId];
        if (font) {
            model.font = font.model
        }
        
        // -------------------------------------------------------
        // alignment
        if (style.alignment) {
            model.alignment = style.alignment.model;
        }
        
        return model;
    },

    // =========================================================================
    // Private Interface
    _addStyle: function(style) {
        var index = this.stylesIndex[style.xml];
        if (index === undefined) {
            this.styles.push(style);
            index = this.stylesIndex[style.xml] = this.styles.length - 1;
        }
        return index;
    },
    
    // =========================================================================
    // Number Formats
    _addNumFmtStr: function(formatCode) {
         // check if default format
        var index = NumFmt.getDefaultFmtId(formatCode);
        if (index !== undefined) return index;
        
        var numFmt = this.numFmtHash[formatCode];
        if (numFmt) return numFmt.id;
        
        // search for an unused index
        index = this.numFmtNextId++;
        return this._addNumFmt(new NumFmt(index, formatCode));
    },
    _addNumFmt: function(numFmt) {
        // called during parse
        this.numFmts.push(numFmt);
        this.numFmtHash[numFmt.id] = numFmt;
        return numFmt.id;
    },
    
    // =========================================================================
    // Fonts
    _addFont: function(font) {
        var fontId = this.fontIndex[font.xml];
        if (fontId === undefined) {
            fontId = this.fontIndex[font.xml] = this.fonts.length;
            this.fonts.push(font);
        }
        return fontId;
    }
    
    // =========================================================================
}

/**
 * Copyright (c) 2015 Guyon Roche
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

var defaultNumFormats = require("./defaultnumformats");
var Font = require("./font");

var StyleManager = module.exports = function() {
    this.cellXfs = [];
    this.cellXfsIndex = {};
    
    this.models = []; // styleId -> model
    
    this.numFmts = [];
    this.numFmtIndex = {}; // formatCode -> numFmtId
    this.numFmtHash = {}; //numFmtId -> numFmt
    this.numFmtNextId = 164; // start custom format ids here
    
    this.fonts = []; // array of font xml
    this.fontIndex = {}; // hash of xml->fontId
    
    // add defaults
    this.addFont(new Font({ size: 11, color: {theme:1}, name: "Calibri", family:2, scheme:"minor"}));
    
    // add default (all zero) style
    this.addXf({numFmtId:0, fontId:0, fillId:0, borderId:0, xfId:0});
}

StyleManager.defaultNumFormats = null;
StyleManager.prototype = {
    addXf: function(xf, force) {
        var key = "" + xf.numFmtId + "." + xf.fontId + "." + xf.fillId + "." + xf.borderId + "." + xf.xfId;
        var index = this.cellXfsIndex[key];
        if ((index === undefined) || force) {
            this.cellXfs.push(xf);
            index = this.cellXfsIndex[key] = this.cellXfs.length - 1;
        }
        return index;
    },
    
    // =========================================================================
    get defaultNumFormats() {
        if (!StyleManager.defaultNumFormats) {
            var dnfs = StyleManager.defaultNumFormats = {};
            _.each(defaultNumFormats, function(dnf, id) {
                if (dnf.f) {
                    dnfs[dnf.f] = id;
                }
                // at some point, add the other cultures here...
            });
        }
        return StyleManager.defaultNumFormats;
    },
    addNumFmtStr: function(formatCode) {
        // called during write()
        var index = this.defaultNumFormats[formatCode];
        if (index !== undefined) return index;
        
        index = this.numFmtIndex[formatCode];
        if (index !== undefined) return index;
        
        // search for an unused index
        index = this.numFmtNextId++;
        var numFmt = {
            numFmtId: index,
            formatCode: formatCode
        };
        this.numFmts.push(numFmt);
        this.numFmtIndex[formatCode] = index;
        this.numFmtHash[index] = numFmt;
        
        return index;
    },
    addNumFmt: function(numFmt) {
        // called during parse
        this.numFmts.push(numFmt);
        this.numFmtIndex[numFmt.formatCode] = numFmt.numFmt;
        this.numFmtHash[numFmt.numFmtId] = numFmt;
        return numFmt.numFmt;
    },
    
    addFont: function(font) {
        var fontId = this.fontIndex[font.xml];
        if (fontId === undefined) {
            fontId = this.fontIndex[font.xml] = this.fonts.length;
            this.fonts.push(font);
        }
        return fontId;
    },
    
    // =========================================================================
    getStyleModel: function(id) {
        var model = this.models[id];
        if (model) return model;
        
        var xf = this.cellXfs[id];
        if (!xf) return null;
        
        var model = this.models[id] = {};
        
        // -------------------------------------------------------
        // number format
        var numFmt = (this.numFmtHash[xf.numFmtId] && this.numFmtHash[xf.numFmtId].formatCode) ||
                    (defaultNumFormats[xf.numFmtId] && defaultNumFormats[xf.numFmtId].f);
        if (numFmt) {
            model.numFmt = numFmt;
        }
        
        // -------------------------------------------------------
        // font
        var font = this.fonts[xf.fontId];
        if (font) {
            model.font = font.model
        }
        
        return model;
    }
}

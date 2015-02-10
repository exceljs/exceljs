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

// Alignment encapsulates translation from style.alignment model to/from xlsx
var Alignment = module.exports = function(model) {
    if (model) {
        this.horizontal = this.validHorizontal(model.horizontal);
        this.vertical = this.validVertical(model.vertical);
        this.wrapText = this.validWrapText(model.wrapText);
        this.indent = this.validIndent(model.indent);
        this.textRotation = this.validTextRotation(model.textRotation);
    }
}

Alignment.prototype = {
    get model() {
        var model = {};
        var hasValues = false;
        if (this.horizontal) {
            hasValues = model.horizontal = this.horizontal;
        }
        if (this.vertical) {
            hasValues = model.vertical = this.vertical;
        }
        if (this.wrapText) {
            hasValues = model.wrapText = this.wrapText;
        }
        if (this.indent) {
            hasValues = model.indent = this.indent;
        }
        if (this.textRotation) {
            hasValues = model.textRotation = this.textRotation;
        }
        return hasValues ? model : null;
    },
    
    get xml() {
        // return string containing the <alignment> definition
        // <alignment horizontal="" vertical="" wrapText="" indent="" textRotation=""/>
        if (this._xml === undefined) {
            var xml = [];
            var isValid = false;
            // build xf node
            xml.push('<alignment');
            if (this.horizontal) {
                xml.push(' horizontal="' + this.horizontal + '"');
                isValid = true;
            }
            if (this.vertical) {
                xml.push(' vertical="' + this.vertical + '"');
                isValid = true;
            }
            if (this.wrapText) {
                xml.push(' wrapText="1"');
                isValid = true;
            }
            if (this.indent && (this.indent > 0)) {
                xml.push(' indent="' + this.indent + '"');
                isValid = true;
            }
            if (this.textRotation) {
                var value;
                if (this.textRotation == "vertical") {
                    value = 255;
                } else {
                    var tr = Math.round(this.textRotation);
                    if ((tr >= 0) && (tr <= 90)) {
                        value = tr;
                    } else if ((tr < 0) && (tr >= -90)) {
                        value = 90 - tr;
                    }
                }
                if (value !== undefined) {
                    xml.push(' textRotation="' + value + '"');
                    isValid = true;
                }
            }
            
            xml.push('/>');
            this._xml = isValid ? xml.join('') : null;
        }
        return this._xml;
    },
    
    parse: function(node) {
        // used during sax parsing of xml to build alignment object
        this.horizontal = node.attributes.horizontal;
        this.vertical = node.attributes.vertical;
        this.wrapText = node.attributes.wrapText ? true : false;
        this.indent = parseInt(node.attributes.indent);
        
        var tr = parseInt(node.attributes.textRotation);
        if (tr !== undefined) {
            if (tr == 255) {
                this.textRotation = "vertical";
            } else if((tr >= 0) && (tr <= 90)) {
                this.textRotation = tr;
            } else if ((tr > 90) && (tr <= 180)) {
                this.textRotation = 90 - tr;
            }
        }
    },
    
    validHorizontal: function(value) {
        switch(value) {
            case "left":
            case "center":
            case "right":
                return value;
            default:
                return undefined;
        }
    },
    validVertical: function(value) {
        switch(value) {
            case "top":
            case "middle":
            case "bottom":
                return value;
            default:
                return undefined;
        }
    },
    validWrapText: function(value) {
        return value ? true : undefined;
    },
    validIndent: function(value) {
        value = parseInt(value);
        return value > 0 ? value : undefined;
    },
    validTextRotation: function(value) {
        switch(value) {
            case "vertical":
                return value;
            default:
                value = parseInt(value);
                return (value >= -90) && (value <= 90) ? value : undefined;
        }
    }
}

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

// Font encapsulates translation from font model to xlsx
var Font = module.exports = function(model) {
    this.model = model || {};
}

Font.prototype = {
    
    get xml() {
        // return string containing the <font> definition
        if (!this._xml) {
            var xml = [];
            var apply = function(xml, tag, attr, value) {
                if (value) {
                    if (value === true) {
                        xml.push('<' + tag + '/>');
                    } else {
                        xml.push('<' + tag + ' ' + attr + '="' + value + '"/>')
                    }
                }
            };
            
            xml.push("<font>");
            
            apply(xml, "b", "val", this.model.bold);
            apply(xml, "i", "val", this.model.italic);
            apply(xml, "u", "val", this.model.underline);
            apply(xml, "strike", "val", this.model.strike);
            apply(xml, "outline", "val", this.model.outline);
            apply(xml, "sz", "val", this.model.size);
            apply(xml, "color", "rgb", this.model.color && this.model.color.argb);
            apply(xml, "color", "theme", this.model.color && this.model.color.theme);
            apply(xml, "name", "val", this.model.name);
            apply(xml, "family", "val", this.model.family);
            apply(xml, "scheme", "val", this.model.scheme);
            apply(xml, "charset", "val", this.model.charset);
            apply(xml, "shadow", "val", this.model.shadow);
            apply(xml, "condense", "val", this.model.condense);
            apply(xml, "extend", "val", this.model.extend);
            
            xml.push("</font>");
            this._xml = xml.join('');
        }
        return this._xml;
    },
    
    parse: function(node) {
        // used during sax parsing of xml to build font object
        switch (node.name) {
            case "b":
                this.model.bold = true;
                break;
            case "i":
                this.model.italic = true;
                break;
            case "u":
                this.model.underline = node.attributes.val || true;
                break;
            case "strike":
                this.model.strike = true;
                break;
            case "outline":
                this.model.outline = true;
                break;
            case "sz":
                this.model.size = parseFloat(node.attributes.val);
                break;
            case "color":
                var c = this.model.color = {};
                if (node.attributes.rgb) {
                    c.argb = node.attributes.rgb;
                } else if (node.attributes.theme) {
                    c.theme = parseInt(node.attributes.theme);
                } else {
                    // it was a mistake!
                    delete this.model.color;
                }
                break;
            case "name":
                this.model.name = node.attributes.val;
                break;
            case "family":
                this.model.family = parseInt(node.attributes.val);
                break;
            case "scheme":
                this.model.scheme = node.attributes.val;
                break;
            case "charset":
                this.model.charset = parseInt(node.attributes.val);
                break;
            case "shadow":
                this.model.shadow = true;
                break;
            case "condense":
                this.model.condense = true;
                break;
            case "extend":
                this.model.extend = true;
                break;
        }
    }
}

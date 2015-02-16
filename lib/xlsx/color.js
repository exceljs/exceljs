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

// Color encapsulates translation from color model to/from xlsx
var Color = module.exports = function(model) {
    if (model) {
        if (model.argb) {
            this.argb = model.argb;
        } else if (model.theme) {
            this.theme = model.theme;
            if (model.tint) { this.tint = model.tint};
        } else {
            this.auto = 1;
        }
    } else {
        this.auto = 1;
    }
}

Color.prototype = {
    get model() {
        if (this.argb) {
            return { argb: this.argb };
        } else if (this.theme) {
            var model = {theme: this.theme};
            if (this.tint) { model.tint = this.tint; }
            return model;
        }
        return null;
    },
    
    get xml() {
        // return string containing the <alignment> definition
        // <color rgb="" or theme="" tint="" or auto="1"/>
        if (this._xml === undefined) {
            if (this.argb) {
                this._xml = '<color rgb="' + this.argb + '"/>';
            } else if (this.theme) {
                if (this.tint) {
                    this._xml = '<color theme="' + this.theme + '" + tint="' + this.tint + '"/>';                    
                } else {
                    this._xml = '<color theme="' + this.theme + '"/>';
                }
            } else {
                this._xml = '<color auto="1"/>';
            }
        }
        return this._xml;
    },
    
    parse: function(node) {
        if (node.attributes.rgb) {
            this.argb = node.attributes.rgb;
        } else if (node.attributes.theme) {
            this.theme = parseInt(node.attributes.theme);
            if (node.attributes.tint) {
                this.tint = parseFloat(node.attributes.tint);
            }
        } else {
            this.auto = true;
        }
    }
}

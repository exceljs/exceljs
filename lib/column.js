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

// Column defines the column properties for 1 column.
// This includes header rows, widths, key, (style), etc.
// Worksheet will condense the columns as appropriate during serialization
var Column = module.exports = function(defn) {
    if (defn) {
        this.headers = defn.header ? [defn.header] : defn.headers || [];
        this.key = defn.key;
        this.width = defn.width;
    }
}

Column.prototype = {
    toString: function() {
        return JSON.stringify({
            key: this.key,
            width: this.width,
            headers: this.headers.length ? this.headers : undefined
        });
    },
    equivalentTo: function(other) {
        return
            (this.width == other.width) &&
            (this.style == other.style);
    },
    get isDefault() {
        return !this.width || (this.width == 8);
    },
    get headerCount() {
        return this.headers.length;
    }
}

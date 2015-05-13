/**
 * Copyright (c) 2014 Guyon Roche
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

var colCache = require("./col-cache");

// used by worksheet to calculate sheet dimensions
var Dimensions = module.exports = function() {
    this.decode(arguments);
}

Dimensions.prototype = {
    decode: function(argv) {
        switch (argv.length) {
            case 4: // [t,l,b,r]
                this.model = {
                    top: Math.min(argv[0], argv[2]),
                    left: Math.min(argv[1], argv[3]),
                    bottom: Math.max(argv[0], argv[2]),
                    right: Math.max(argv[1], argv[3])
                };
                break;
                
            case 2: // [tl, br]
                var tl = colCache.decodeAddress(argv[0]);
                var br = colCache.decodeAddress(argv[1]);
                this.model = {
                    top: Math.min(tl.row, br.row),
                    left: Math.min(tl.col, br.col),
                    bottom: Math.max(tl.row, br.row),
                    right: Math.max(tl.col, br.col)
                };
                break;
                
            case 1: // tl:br
                var value = argv[0];
                if (value instanceof Dimensions) { // copy constructor
                    this.model = {
                        top: value.model.top,
                        left: value.model.left,
                        bottom: value.model.bottom,
                        right: value.model.right
                    };
                } else if (value instanceof Array) { // an arguments array
                    this.decode(value);
                } else if (value.top && value.left && value.bottom && value.right) { // a model
                    this.model = value;
                } else { // tl:br string
                    var tlbr = colCache.decode(value);
                    if (tlbr.top) {
                        this.model = {
                            top: tlbr.top,
                            left: tlbr.left,
                            bottom: tlbr.bottom,
                            right: tlbr.right
                        };
                    } else {
                        this.model = {
                            top: tlbr.row,
                            left: tlbr.col,
                            bottom: tlbr.row,
                            right: tlbr.col
                        };
                    }
                }
                break;
            
            case 0:
                this.model = {
                    top: 0,
                    left: 0,
                    bottom: 0,
                    right: 0
                };
                break;
            
            default:
                throw new Error("Invalid number of arguments to _getDimensions() - " + argv.length);
        }
    },
    
    get top() { return this.model.top || 1; },
    set top(value) { return this.model.top = value; },
    get left() { return this.model.left || 1; },
    set left(value) { return this.model.left = value; },
    get bottom() { return this.model.bottom || 1; },
    set bottom(value) { return this.model.bottom = value; },
    get right() { return this.model.right || 1; },
    set right(value) { return this.model.right = value; },
    
    expand: function(top, left, bottom, right) {
        if (!this.model.top || (top < this.top)) this.top = top;
        if (!this.model.left || (left < this.left)) this.left = left;
        if (!this.model.bottom || (bottom > this.bottom)) this.bottom = bottom;
        if (!this.model.right || (right > this.right)) this.right = right;
    },
    
    get tl() {
        return colCache.n2l(this.left || 1) + (this.top || 1);
    },
    get br() {
        return colCache.n2l(this.right || 1) + (this.bottom || 1);
    },
    
    get range() {
        return this.tl + ":" + this.br;
    },
    toString: function() {
        return this.range;
    },
    
    intersects: function(other) {
        if (other.bottom < this.top) return false;
        if (other.top > this.bottom) return false;
        if (other.right < this.left) return false;
        if (other.left > this.right) return false;
        return true;
    }
}
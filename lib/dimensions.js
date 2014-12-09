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

var colCache = require("./colcache");

// used by worksheet to calculate sheet dimensions
var Dimensions = module.exports = function() {
    this.decode(arguments);
}

Dimensions.prototype = {
    decode: function(argv) {
        switch (argv.length) {
            case 4: // [t,l,b,r]
                this.top = Math.min(argv[0], argv[2]);
                this.left = Math.min(argv[1], argv[3]);
                this.bottom = Math.max(argv[2], argv[2]);
                this.right = Math.max(argv[1], argv[3]);
                break;
                
            case 2: // [tl, br]
                var tl = colCache.decodeAddress(argv[0]);
                var br = colCache.decodeAddress(argv[1]);
                this.top = Math.min(tl.row, br.row);
                this.left = Math.min(tl.col, br.col);
                this.bottom = Math.max(tl.row, br.row);
                this.right = Math.max(tl.col, br.col);
                break;
                
            case 1: // tl:br
                var value = argv[0];
                if (value instanceof Array) { // an arguments array
                    this.decode(value);
                } else if (value.top && value.left) { // a model
                    this.model = value;
                } else { // tl:br string
                    var tlbr = colCache.decode(value);
                    this.top = tlbr.top;
                    this.left = tlbr.left;
                    this.bottom = tlbr.bottom;
                    this.right = tlbr.right;
                }
                break;
            
            case 0:
                this.top = 0;
                this.left = 0;
                this.bottom = 0;
                this.right = 0;
                break;
            
            default:
                throw new Error("Invalid number of arguments to _getDimensions() - " + argv.length);
        }
    },
    
    get model() {
        return {
            top: this.top,
            left: this.left,
            bottom: this.bottom,
            right: this.right
        };
    },
    set model(value) {
        this.top = value.top;
        this.left = value.left;
        this.bottom = value.bottom;
        this.right = value.right;
    },
    expand: function(top, left, bottom, right) {
        if (!this.top || (top < this.top)) this.top = top;
        if (!this.left || (left < this.left)) this.left = left;
        if (!this.bottom || (bottom > this.bottom)) this.bottom = bottom;
        if (!this.right || (right > this.right)) this.right = right;
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
    }
}
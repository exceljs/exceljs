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

// =========================================================================
// Column Letter to Number conversion
var colCache = module.exports = {
    _dictionary: ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"],
    _l2n: {},
    _n2l: [],
    _level: function(n) {
        if (n <= 26) return 1;
        if (n <= 26*26) return 2;
        return 3;
    },
    _fill: function(level) {
        var n = 1;
        if (level >= 1) {
            //console.log("Level 1")
            while (n <= 26) {
                var c = this._dictionary[n-1];
                this._n2l[n] = c;
                this._l2n[c] = n;
                n++;
            }
            //console.log(this._n2l[1] + "-" + this._n2l[26]);
        }
        if (level >= 2) {
            //console.log("Level 2")
            while (n <= 26 + 26 * 26) {
                var v = n-(26+1);
                var l1 = v % 26;
                var l2 = Math.floor(v / 26);
                var c = this._dictionary[l2] + this._dictionary[l1];
                this._n2l[n] = c;
                this._l2n[c] = n;
                n++;
            }
            //console.log(this._n2l[27] + "-" + this._n2l[26+26*26]);
        }
        if (level >= 3) {
            while (n <= 16384) {
                var v = n-(26*26+26+1);
                var l1 = v % 26;
                var l2 = Math.floor(v / 26) % 26;
                var l3 = Math.floor(v / (26*26));
                var c = this._dictionary[l3] + this._dictionary[l2] + this._dictionary[l1];
                this._n2l[n] = c;
                this._l2n[c] = n;
                n++;
            }
        }
    },
    l2n: function(l) {
        if (!this._l2n[l]) {
            this._fill(l.length);
        }
        if (!this._l2n[l]) {
            throw new Error("Out of bounds. Invalid column letter: " + l);
        }
        return this._l2n[l];
    },
    n2l: function(n) {
        if ((n < 1) || (n > 16384)) {
            throw new Error("Out of bounds. Excel supports columns from 1 to 16384");
        }
        if (!this._n2l[n]) { 
            this._fill(this._level(n));
        }
        return this._n2l[n];
    },
    
    // =========================================================================
    // Address processing
    _hash: {},
    
    // check if value looks like an address
    validateAddress: function(value) {
        if (!value.match(/^[A-Z]+\d+$/)) {
            throw new Error("Invalid Address: " + value);
        }
        return true;
    },
    
    // convert address string into structure
    decodeAddress: function(value) {
        var addr = this._hash[value];
        if (addr) return addr;
        return this._hash[value] = {
            address: value,
            col: this.l2n(value.replace(/\d+/, "")),
            row: parseInt(value.replace(/[A-Z]+/, ""))
        };        
    },
    
    // convert r,c into structure (if only 1 arg, assume r is address string)
    getAddress: function(r,c) {
        if (c) {
            var address = this.n2l(c) + r;
            return this.decodeAddress(address);
        } else {
            return this.decodeAddress(r);
        }
    },
    
    // convert [address], [tl:br] into address structures
    decode: function(value) {
        var parts = value.split(':');
        if (parts.length == 2) {
            var tl = this.decodeAddress(parts[0]);
            var br = this.decodeAddress(parts[1]);
            var result = {
                top: Math.min(tl.row, br.row),
                left: Math.min(tl.col, br.col),
                bottom: Math.max(tl.row, br.row),
                right: Math.max(tl.col, br.col)
            }
            // reconstruct tl, br and dimensions
            result.tl = this.n2l(result.left) + result.top;
            result.br = this.n2l(result.right) + result.bottom;
            result.dimensions = result.tl + ":" + result.br;
            return result;
        } else {
            return this.decodeAddress(value);
        }
    },
    
    // convert row,col into address string
    encodeAddress: function(row, col) {
        return colCache.n2l(col) + row;
    },
    
    // convert row,col into string address or t,l,b,r into range
    encode: function() {
        switch(arguments.length) {
            case 2:
                return colCache.encodeAddress(arguments[0], arguments[1]);
            case 4:
                return colCache.encodeAddress(arguments[0], arguments[1]) + ":" + colCache.encodeAddress(arguments[2], arguments[3]);
            default:
                throw new Error("Can only encode with 2 or 4 arguments");
        }
    }
};


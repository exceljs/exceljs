/**
 * Copyright (c) 2016 Guyon Roche
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

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var IntegerXform = require('../simple/integer-xform');

var CellPositionXform = module.exports = function(options) {
  this.tag = options.tag;
  this.map = {
    'xdr:col': new IntegerXform({tag: 'xdr:col', zero: true}),
    'xdr:colOff': new IntegerXform({tag: 'xdr:colOff', zero: true}),
    'xdr:row': new IntegerXform({tag: 'xdr:row', zero: true}),
    'xdr:rowOff': new IntegerXform({tag: 'xdr:rowOff', zero: true})
  };
};

utils.inherits(CellPositionXform, BaseXform, {

  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag);

    var col = Math.floor(model.col);
    var colOff = Math.floor((model.col - col) * 640000);
    this.map['xdr:col'].render(xmlStream, col);
    this.map['xdr:colOff'].render(xmlStream, colOff);

    var row = Math.floor(model.row);
    var rowOff = Math.floor((model.row - row) * 180000);
    this.map['xdr:row'].render(xmlStream, row);
    this.map['xdr:rowOff'].render(xmlStream, rowOff);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  },

  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch(name) {
      case this.tag:
        this.model = {
          col: this.map['xdr:col'].model + this.map['xdr:colOff'].model / 640000,
          row: this.map['xdr:row'].model + this.map['xdr:rowOff'].model / 180000
        };
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
});

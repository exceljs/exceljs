/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var IntegerXform = require('../simple/integer-xform');

var CellPositionXform = module.exports = function (options) {
  this.tag = options.tag;
  this.map = {
    'xdr:col': new IntegerXform({ tag: 'xdr:col', zero: true }),
    'xdr:colOff': new IntegerXform({ tag: 'xdr:colOff', zero: true }),
    'xdr:row': new IntegerXform({ tag: 'xdr:row', zero: true }),
    'xdr:rowOff': new IntegerXform({ tag: 'xdr:rowOff', zero: true })
  };
};

utils.inherits(CellPositionXform, BaseXform, {

  render: function render(xmlStream, model) {
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

  parseOpen: function parseOpen(node) {
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

  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
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
//# sourceMappingURL=cell-position-xform.js.map

/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var IntegerXform = require('../simple/integer-xform');

var CellPositionXform = module.exports = function(options) {
  this.tag = options.tag;
  this.map = {
    'col': new IntegerXform({tag: 'col', ns: 'xdr', zero: true}),
    'colOff': new IntegerXform({tag: 'colOff', ns: 'xdr', zero: true}),
    'row': new IntegerXform({tag: 'row', ns: 'xdr', zero: true}),
    'rowOff': new IntegerXform({tag: 'rowOff', ns: 'xdr', zero: true})
  };
};

utils.inherits(CellPositionXform, BaseXform, {

  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag);

    var col = Math.floor(model.col);
    var colOff = Math.floor((model.col - col) * 640000);
    this.map['col'].render(xmlStream, col);
    this.map['colOff'].render(xmlStream, colOff);

    var row = Math.floor(model.row);
    var rowOff = Math.floor((model.row - row) * 180000);
    this.map['row'].render(xmlStream, row);
    this.map['rowOff'].render(xmlStream, rowOff);

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
    switch (name) {
      case this.tag:
        this.model = {
          col: this.map['col'].model + (this.map['colOff'].model / 640000),
          row: this.map['row'].model + (this.map['rowOff'].model / 180000)
        };
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
});

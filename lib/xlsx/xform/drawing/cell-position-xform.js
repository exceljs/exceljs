'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var IntegerXform = require('../simple/integer-xform');
var Anchor = require('../../../doc/anchor');

var CellPositionXform = module.exports = function(options) {
  this.tag = options.tag;
  this.map = {
    'xdr:col': new IntegerXform({tag: 'xdr:col', zero: true}),
    'xdr:colOff': new IntegerXform({tag: 'xdr:colOff', zero: true}),
    'xdr:row': new IntegerXform({tag: 'xdr:row', zero: true}),
    'xdr:rowOff': new IntegerXform({tag: 'xdr:rowOff', zero: true})
  };
};
CellPositionXform.buildModel = function (model) {
  return new Anchor(model);
};


utils.inherits(CellPositionXform, BaseXform, {

  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:col'].render(xmlStream, model.nativeCol);
    this.map['xdr:colOff'].render(xmlStream, model.nativeColOff);

    this.map['xdr:row'].render(xmlStream, model.nativeRow);
    this.map['xdr:rowOff'].render(xmlStream, model.nativeRowOff);

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
        this.model = CellPositionXform.buildModel({
          nativeCol: this.map['xdr:col'].model,
          nativeColOff: this.map['xdr:colOff'].model,
          nativeRow: this.map['xdr:row'].model,
          nativeRowOff: this.map['xdr:rowOff'].model
        });
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
});

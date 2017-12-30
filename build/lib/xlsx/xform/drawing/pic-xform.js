/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');

var BlipFillXform = require('./blip-fill-xform');
var NvPicPrXform = require('./nv-pic-pr-xform');

var spPrJSON = require('./sp-pr');

var PicXform = module.exports = function () {
  this.map = {
    'xdr:nvPicPr': new NvPicPrXform(),
    'xdr:blipFill': new BlipFillXform(),
    'xdr:spPr': new StaticXform(spPrJSON)
  };
};

utils.inherits(PicXform, BaseXform, {
  get tag() {
    return 'xdr:pic';
  },

  prepare: function prepare(model, options) {
    model.index = options.index + 1;
  },

  render: function render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:nvPicPr'].render(xmlStream, model);
    this.map['xdr:blipFill'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);

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

  parseText: function parseText() {},

  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.mergeModel(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
});
//# sourceMappingURL=pic-xform.js.map

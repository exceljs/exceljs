/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

// Color encapsulates translation from color model to/from xlsx
var ColorXform = module.exports = function (name) {
  // this.name controls the xm node name
  this.name = name || 'color';
};

utils.inherits(ColorXform, BaseXform, {

  get tag() {
    return this.name;
  },

  render: function render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.name);
      if (model.argb) {
        xmlStream.addAttribute('rgb', model.argb);
      } else if (model.theme !== undefined) {
        xmlStream.addAttribute('theme', model.theme);
        if (model.tint !== undefined) {
          xmlStream.addAttribute('tint', model.tint);
        }
      } else if (model.indexed !== undefined) {
        xmlStream.addAttribute('indexed', model.indexed);
      } else {
        xmlStream.addAttribute('auto', '1');
      }
      xmlStream.closeNode();
      return true;
    }
    return false;
  },

  parseOpen: function parseOpen(node) {
    if (node.name === this.name) {
      if (node.attributes.rgb) {
        this.model = { argb: node.attributes.rgb };
      } else if (node.attributes.theme) {
        this.model = { theme: parseInt(node.attributes.theme, 10) };
        if (node.attributes.tint) {
          this.model.tint = parseFloat(node.attributes.tint);
        }
      } else if (node.attributes.indexed) {
        this.model = { indexed: parseInt(node.attributes.indexed, 10) };
      } else {
        this.model = undefined;
      }
      return true;
    }
    return false;
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=color-xform.js.map

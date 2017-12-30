/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var FloatXform = module.exports = function (options) {
  this.tag = options.tag;
  this.attr = options.attr;
  this.attrs = options.attrs;
};

utils.inherits(FloatXform, BaseXform, {

  render: function render(xmlStream, model) {
    if (model !== undefined) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, model);
      } else {
        xmlStream.writeText(model);
      }
      xmlStream.closeNode();
    }
  },

  parseOpen: function parseOpen(node) {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = parseFloat(node.attributes[this.attr]);
      } else {
        this.text = [];
      }
    }
  },
  parseText: function parseText(text) {
    if (!this.attr) {
      this.text.push(text);
    }
  },
  parseClose: function parseClose() {
    if (!this.attr) {
      this.model = parseFloat(this.text.join(''));
    }
    return false;
  }
});
//# sourceMappingURL=float-xform.js.map

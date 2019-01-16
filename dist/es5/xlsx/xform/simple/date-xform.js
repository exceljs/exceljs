/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var DateXform = module.exports = function (options) {
  this.tag = options.tag;
  this.attr = options.attr;
  this.attrs = options.attrs;
  this._format = options.format || function (dt) {
    try {
      if (isNaN(dt.getTime())) return '';
      return dt.toISOString();
    } catch (e) {
      return '';
    }
  };
  this._parse = options.parse || function (str) {
    return new Date(str);
  };
};

utils.inherits(DateXform, BaseXform, {

  render: function render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.tag);
      if (this.attrs) {
        xmlStream.addAttributes(this.attrs);
      }
      if (this.attr) {
        xmlStream.addAttribute(this.attr, this._format(model));
      } else {
        xmlStream.writeText(this._format(model));
      }
      xmlStream.closeNode();
    }
  },

  parseOpen: function parseOpen(node) {
    if (node.name === this.tag) {
      if (this.attr) {
        this.model = this._parse(node.attributes[this.attr]);
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
      this.model = this._parse(this.text.join(''));
    }
    return false;
  }
});
//# sourceMappingURL=date-xform.js.map

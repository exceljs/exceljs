/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

//   <t xml:space="preserve"> is </t>

var TextXform = module.exports = function () {};

utils.inherits(TextXform, BaseXform, {

  get tag() {
    return 't';
  },

  render: function render(xmlStream, model) {
    xmlStream.openNode('t');
    if (model[0] === ' ' || model[model.length - 1] === ' ') {
      xmlStream.addAttribute('xml:space', 'preserve');
    }
    xmlStream.writeText(model);
    xmlStream.closeNode();
  },

  get model() {
    return this._text.join('').replace(/_x([0-9A-F]{4})_/g, function ($0, $1) {
      return String.fromCharCode(parseInt($1, 16));
    });
  },

  parseOpen: function parseOpen(node) {
    switch (node.name) {
      case 't':
        this._text = [];
        return true;
      default:
        return false;
    }
  },
  parseText: function parseText(text) {
    this._text.push(text);
  },
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=text-xform.js.map

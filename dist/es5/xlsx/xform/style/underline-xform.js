/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var UnderlineXform = module.exports = function (model) {
  this.model = model;
};

UnderlineXform.Attributes = {
  single: {},
  double: { val: 'double' },
  singleAccounting: { val: 'singleAccounting' },
  doubleAccounting: { val: 'doubleAccounting' }
};

utils.inherits(UnderlineXform, BaseXform, {

  get tag() {
    return 'u';
  },

  render: function render(xmlStream, model) {
    model = model || this.model;

    if (model === true) {
      xmlStream.leafNode('u');
    } else {
      var attr = UnderlineXform.Attributes[model];
      if (attr) {
        xmlStream.leafNode('u', attr);
      }
    }
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'u') {
      this.model = node.attributes.val || true;
    }
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=underline-xform.js.map

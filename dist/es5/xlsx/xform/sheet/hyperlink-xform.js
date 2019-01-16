/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var HyperlinkXform = module.exports = function () {};

utils.inherits(HyperlinkXform, BaseXform, {

  get tag() {
    return 'hyperlink';
  },

  render: function render(xmlStream, model) {
    xmlStream.leafNode('hyperlink', {
      ref: model.address,
      'r:id': model.rId
    });
  },

  parseOpen: function parseOpen(node) {
    if (node.name === 'hyperlink') {
      this.model = {
        address: node.attributes.ref,
        rId: node.attributes['r:id']
      };
      return true;
    }
    return false;
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=hyperlink-xform.js.map

/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const BooleanXform = (module.exports = function(options) {
  this.tag = options.tag;
  this.attr = options.attr;
});

utils.inherits(BooleanXform, BaseXform, {
  render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  },

  parseOpen(node) {
    if (node.name === this.tag) {
      this.model = true;
    }
  },
  parseText() {},
  parseClose() {
    return false;
  },
});

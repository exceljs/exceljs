/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const DimensionXform = (module.exports = function() {});

utils.inherits(DimensionXform, BaseXform, {
  get tag() {
    return 'dimension';
  },

  render(xmlStream, model) {
    if (model) {
      xmlStream.leafNode('dimension', { ref: model });
    }
  },

  parseOpen(node) {
    if (node.name === 'dimension') {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  },
  parseText() {},
  parseClose() {
    return false;
  },
});

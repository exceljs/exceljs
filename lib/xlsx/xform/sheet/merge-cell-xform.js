/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const MergeCellXform = (module.exports = function() {});

utils.inherits(MergeCellXform, BaseXform, {
  get tag() {
    return 'mergeCell';
  },

  render(xmlStream, model) {
    xmlStream.leafNode('mergeCell', { ref: model });
  },

  parseOpen(node) {
    if (node.name === 'mergeCell') {
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

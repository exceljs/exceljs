/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const PageBreaksXform = (module.exports = function() {});

utils.inherits(PageBreaksXform, BaseXform, {
  get tag() {
    return 'brk';
  },

  render(xmlStream, model) {
    xmlStream.leafNode('brk', model);
  },

  parseOpen(node) {
    if (node.name === 'brk') {
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

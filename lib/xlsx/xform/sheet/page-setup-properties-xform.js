/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const PageSetupPropertiesXform = (module.exports = function() {});

utils.inherits(PageSetupPropertiesXform, BaseXform, {
  get tag() {
    return 'pageSetUpPr';
  },

  render(xmlStream, model) {
    if (model && model.fitToPage) {
      xmlStream.leafNode(this.tag, {
        fitToPage: model.fitToPage ? '1' : undefined,
      });
      return true;
    }
    return false;
  },

  parseOpen(node) {
    if (node.name === this.tag) {
      this.model = {
        fitToPage: node.attributes.fitToPage === '1',
      };
      return true;
    }
    return false;
  },
  parseText() {},
  parseClose() {
    return false;
  },
});

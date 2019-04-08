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

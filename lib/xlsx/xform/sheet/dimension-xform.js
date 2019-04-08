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

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const DrawingXform = (module.exports = function() {});

utils.inherits(DrawingXform, BaseXform, {
  get tag() {
    return 'drawing';
  },

  render(xmlStream, model) {
    if (model) {
      xmlStream.leafNode(this.tag, { 'r:id': model.rId });
    }
  },

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          rId: node.attributes['r:id'],
        };
        return true;
      default:
        return false;
    }
  },
  parseText() {},
  parseClose() {
    return false;
  },
});

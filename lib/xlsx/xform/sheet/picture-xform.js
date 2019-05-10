'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const PictureXform = (module.exports = function() {});

utils.inherits(PictureXform, BaseXform, {
  get tag() {
    return 'picture';
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

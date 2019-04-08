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

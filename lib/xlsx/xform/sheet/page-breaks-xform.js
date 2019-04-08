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

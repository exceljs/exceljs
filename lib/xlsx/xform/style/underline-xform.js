'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const UnderlineXform = (module.exports = function(model) {
  this.model = model;
});

UnderlineXform.Attributes = {
  single: {},
  double: { val: 'double' },
  singleAccounting: { val: 'singleAccounting' },
  doubleAccounting: { val: 'doubleAccounting' },
};

utils.inherits(UnderlineXform, BaseXform, {
  get tag() {
    return 'u';
  },

  render(xmlStream, model) {
    model = model || this.model;

    if (model === true) {
      xmlStream.leafNode('u');
    } else {
      const attr = UnderlineXform.Attributes[model];
      if (attr) {
        xmlStream.leafNode('u', attr);
      }
    }
  },

  parseOpen(node) {
    if (node.name === 'u') {
      this.model = node.attributes.val || true;
    }
  },
  parseText() {},
  parseClose() {
    return false;
  },
});

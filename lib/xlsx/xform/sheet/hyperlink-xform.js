'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const HyperlinkXform = (module.exports = function() {});

utils.inherits(HyperlinkXform, BaseXform, {
  get tag() {
    return 'hyperlink';
  },

  render(xmlStream, model) {
    xmlStream.leafNode(
      'hyperlink',
      Object.assign(
        {
          ref: model.address,
          'r:id': model.rId,
        },
        model.tooltip ? { tooltip: model.tooltip } : {}
      )
    );
  },

  parseOpen(node) {
    if (node.name === 'hyperlink') {
      this.model = Object.assign(
        {
          address: node.attributes.ref,
          rId: node.attributes['r:id'],
        },
        node.attributes.tooltip ? { tooltip: node.attributes.tooltip } : {}
      );
      return true;
    }
    return false;
  },
  parseText() {},
  parseClose() {
    return false;
  },
});

'use strict';

const utils = require('../../../utils/utils');
const colCache = require('../../../utils/col-cache');
const BaseXform = require('../base-xform');

const AutoFilterXform = (module.exports = function() {});

utils.inherits(AutoFilterXform, BaseXform, {
  get tag() {
    return 'autoFilter';
  },

  render(xmlStream, model) {
    if (model) {
      if (typeof model === 'string') {
        // assume range
        xmlStream.leafNode('autoFilter', { ref: model });
      } else {
        const getAddress = function(addr) {
          if (typeof addr === 'string') {
            return addr;
          }
          return colCache.getAddress(addr.row, addr.column).address;
        };

        const firstAddress = getAddress(model.from);
        const secondAddress = getAddress(model.to);
        if (firstAddress && secondAddress) {
          xmlStream.leafNode('autoFilter', { ref: `${firstAddress}:${secondAddress}` });
        }
      }
    }
  },

  parseOpen(node) {
    if (node.name === 'autoFilter') {
      this.model = node.attributes.ref;
    }
  },
});

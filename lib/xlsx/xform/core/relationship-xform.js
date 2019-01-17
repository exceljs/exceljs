/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const RelationshipXform = (module.exports = function() {});

utils.inherits(RelationshipXform, BaseXform, {
  render(xmlStream, model) {
    xmlStream.leafNode('Relationship', model);
  },

  parseOpen(node) {
    switch (node.name) {
      case 'Relationship':
        this.model = node.attributes;
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

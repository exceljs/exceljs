/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var AppTitlesOfPartsXform = module.exports = function () {};

utils.inherits(AppTitlesOfPartsXform, BaseXform, {
  render: function render(xmlStream, model) {
    xmlStream.openNode('TitlesOfParts');
    xmlStream.openNode('vt:vector', { size: model.length, baseType: 'lpstr' });

    model.forEach(function (sheet) {
      xmlStream.leafNode('vt:lpstr', undefined, sheet.name);
    });

    xmlStream.closeNode();
    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    // no parsing
    return node.name === 'TitlesOfParts';
  },
  parseText: function parseText() {},
  parseClose: function parseClose(name) {
    return name !== 'TitlesOfParts';
  }
});
//# sourceMappingURL=app-titles-of-parts-xform.js.map

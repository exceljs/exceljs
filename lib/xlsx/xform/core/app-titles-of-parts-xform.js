'use strict';

const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

const AppTitlesOfPartsXform = (module.exports = function() {});

utils.inherits(AppTitlesOfPartsXform, BaseXform, {
  render(xmlStream, model) {
    xmlStream.openNode('TitlesOfParts');
    xmlStream.openNode('vt:vector', { size: model.length, baseType: 'lpstr' });

    model.forEach(sheet => {
      xmlStream.leafNode('vt:lpstr', undefined, sheet.name);
    });

    xmlStream.closeNode();
    xmlStream.closeNode();
  },

  parseOpen(node) {
    // no parsing
    return node.name === 'TitlesOfParts';
  },
  parseText() {},
  parseClose(name) {
    return name !== 'TitlesOfParts';
  },
});

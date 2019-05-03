'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var AppTitlesOfPartsXform = module.exports = function() {
};

utils.inherits(AppTitlesOfPartsXform, BaseXform, {
  render: function(xmlStream, model) {
    xmlStream.openNode('TitlesOfParts');
    xmlStream.openNode('vt:vector', {size: model.length, baseType: 'lpstr'});

    model.forEach(function(sheet) {
      xmlStream.leafNode('vt:lpstr', undefined, sheet.name);
    });

    xmlStream.closeNode();
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    // no parsing
    return node.name === 'TitlesOfParts';
  },
  parseText: function() {
  },
  parseClose: function(name) {
    return name !== 'TitlesOfParts';
  }
});

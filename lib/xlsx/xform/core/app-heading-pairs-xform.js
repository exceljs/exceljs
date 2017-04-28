/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var AppHeadingPairsXform = module.exports = function() {
};

utils.inherits(AppHeadingPairsXform, BaseXform, {
  render: function(xmlStream, model) {
    xmlStream.openNode('HeadingPairs');
    xmlStream.openNode('vt:vector', {size: 2, baseType: 'variant'});

    xmlStream.openNode('vt:variant');
    xmlStream.leafNode('vt:lpstr', undefined, 'Worksheets');
    xmlStream.closeNode();
    
    xmlStream.openNode('vt:variant');
    xmlStream.leafNode('vt:i4', undefined, model.length);
    xmlStream.closeNode();
    
    xmlStream.closeNode();
    xmlStream.closeNode();
  },
  
  parseOpen: function(node) {
    // no parsing
    return node.name === 'HeadingPairs';
  },
  parseText: function() {
  },
  parseClose: function(name) {
    return name !== 'HeadingPairs';
  }
});

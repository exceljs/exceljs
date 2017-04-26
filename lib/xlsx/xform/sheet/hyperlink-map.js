/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var HyperlinkMap = function(relationships, hyperlinks) {
  var rels = {};
  var map = this.map = {};
  if (relationships && hyperlinks) {
    relationships.forEach(function (relationship) {
      rels[relationship.rId] = relationship.target;
    });

    hyperlinks.forEach(function (hyperlink) {
      map[hyperlink.address] = rels[hyperlink.rId];
    });
  }
};

HyperlinkMap.prototype = {
  getHyperlink: function(address) {
    return this.map[address];
  }
};

module.exports = HyperlinkMap;

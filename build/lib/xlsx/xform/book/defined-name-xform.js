/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var DefinedNamesXform = module.exports = function () {};

utils.inherits(DefinedNamesXform, BaseXform, {
  render: function render(xmlStream, model) {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode('definedName', {
      name: model.name,
      localSheetId: model.localSheetId
    });
    xmlStream.writeText(model.ranges.join(','));
    xmlStream.closeNode();
  },
  parseOpen: function parseOpen(node) {
    switch (node.name) {
      case 'definedName':
        this._parsedName = node.attributes.name;
        this._parsedLocalSheetId = node.attributes.localSheetId;
        this._parsedText = [];
        return true;
      default:
        return false;
    }
  },
  parseText: function parseText(text) {
    this._parsedText.push(text);
  },
  parseClose: function parseClose() {
    this.model = {
      name: this._parsedName,
      ranges: extractRanges(this._parsedText.join(''))
    };
    if (this._parsedLocalSheetId !== undefined) {
      this.model.localSheetId = parseInt(this._parsedLocalSheetId, 10);
    }
    return false;
  }
});

function extractRanges(parsedText) {
  var ranges = [];
  var quotesOpened = false;
  var last = '';
  parsedText.split(',').forEach(function (item) {
    if (!item) {
      return;
    }
    var quotes = (item.match(/'/g) || []).length;

    if (!quotes) {
      if (quotesOpened) {
        last += item + ',';
      } else {
        ranges.push(item);
      }
      return;
    }
    var quotesEven = quotes % 2 === 0;

    if (!quotesOpened && quotesEven) {
      ranges.push(item);
    } else if (quotesOpened && !quotesEven) {
      quotesOpened = false;
      ranges.push(last + item);
      last = '';
    } else {
      quotesOpened = true;
      last += item + ',';
    }
  });
  return ranges;
}
//# sourceMappingURL=defined-name-xform.js.map

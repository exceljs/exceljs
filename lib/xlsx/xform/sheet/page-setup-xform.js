/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

function booleanToXml(model) {
  return model ? '1' : undefined;
}
function pageOrderToXml(model) {
  switch(model) {
    case 'overThenDown': 
      return model;
    default:
      return undefined;
  }
}
function cellCommentsToXml(model) {
  switch(model) {
    case 'atEnd': 
    case 'asDisplyed': 
      return model;
    default:
      return undefined;
  }
}
function errorsToXml(model) {
  switch(model) {
    case 'dash': 
    case 'blank': 
    case 'NA': 
      return model;
    default:
      return undefined;
  }
}
function pageSizeToModel(value) {
  return value !== undefined ? parseInt(value) : undefined;
}

var PageSetupXform = module.exports = function() {
};

utils.inherits(PageSetupXform, BaseXform, {

  get tag() { return 'pageSetup'; },

  render: function(xmlStream, model) {
    if (model) {
      var attributes = {
        paperSize: model.paperSize,
        orientation: model.orientation,
        horizontalDpi: model.horizontalDpi,
        verticalDpi: model.verticalDpi,
        pageOrder: pageOrderToXml(model.pageOrder),
        blackAndWhite: booleanToXml(model.blackAndWhite),
        draft: booleanToXml(model.draft),
        cellComments: cellCommentsToXml(model.cellComments),
        errors: errorsToXml(model.errors),
        scale: model.scale,
        fitToWidth: model.fitToWidth,
        fitToHeight: model.fitToHeight,
        firstPageNumber: model.firstPageNumber,
        useFirstPageNumber: booleanToXml(model.firstPageNumber),
        usePrinterDefaults: booleanToXml(model.usePrinterDefaults),
        copies: model.copies
      };
      if (_.some(attributes, function (value) {
          return value !== undefined;
        })) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  },
  
  parseOpen: function(node) {
    switch(node.name) {
      case this.tag: 
        this.model = {
          paperSize: pageSizeToModel(node.attributes.paperSize),
          orientation: node.attributes.orientation || 'portrait',
          horizontalDpi: parseInt(node.attributes.horizontalDpi || 4294967295),
          verticalDpi: parseInt(node.attributes.verticalDpi || 4294967295),
          pageOrder: node.attributes.pageOrder || 'downThenOver',
          blackAndWhite: node.attributes.blackAndWhite === '1',
          draft: node.attributes.draft === '1',
          cellComments: node.attributes.cellComments || 'None',
          errors: node.attributes.errors || 'displayed',
          scale: parseInt(node.attributes.scale || 100),
          fitToWidth: parseInt(node.attributes.fitToWidth || 1),
          fitToHeight: parseInt(node.attributes.fitToHeight || 1),
          firstPageNumber: parseInt(node.attributes.firstPageNumber || 1),
          useFirstPageNumber: node.attributes.useFirstPageNumber === '1',
          usePrinterDefaults: node.attributes.usePrinterDefaults === '1',
          copies: parseInt(node.attributes.copies || 1)
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});

/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

function booleanToXml(model) {
  return model ? '1' : undefined;
}
function pageOrderToXml(model) {
  switch (model) {
    case 'overThenDown':
      return model;
    default:
      return undefined;
  }
}
function cellCommentsToXml(model) {
  switch (model) {
    case 'atEnd':
    case 'asDisplyed':
      return model;
    default:
      return undefined;
  }
}
function errorsToXml(model) {
  switch (model) {
    case 'dash':
    case 'blank':
    case 'NA':
      return model;
    default:
      return undefined;
  }
}
function pageSizeToModel(value) {
  return value !== undefined ? parseInt(value, 10) : undefined;
}

var PageSetupXform = module.exports = function () {};

utils.inherits(PageSetupXform, BaseXform, {

  get tag() {
    return 'pageSetup';
  },

  render: function render(xmlStream, model) {
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

  parseOpen: function parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          paperSize: pageSizeToModel(node.attributes.paperSize),
          orientation: node.attributes.orientation || 'portrait',
          horizontalDpi: parseInt(node.attributes.horizontalDpi || '4294967295', 10),
          verticalDpi: parseInt(node.attributes.verticalDpi || '4294967295', 10),
          pageOrder: node.attributes.pageOrder || 'downThenOver',
          blackAndWhite: node.attributes.blackAndWhite === '1',
          draft: node.attributes.draft === '1',
          cellComments: node.attributes.cellComments || 'None',
          errors: node.attributes.errors || 'displayed',
          scale: parseInt(node.attributes.scale || '100', 10),
          fitToWidth: parseInt(node.attributes.fitToWidth || '1', 10),
          fitToHeight: parseInt(node.attributes.fitToHeight || '1', 10),
          firstPageNumber: parseInt(node.attributes.firstPageNumber || '1', 10),
          useFirstPageNumber: node.attributes.useFirstPageNumber === '1',
          usePrinterDefaults: node.attributes.usePrinterDefaults === '1',
          copies: parseInt(node.attributes.copies || '1', 10)
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function parseText() {},
  parseClose: function parseClose() {
    return false;
  }
});
//# sourceMappingURL=page-setup-xform.js.map

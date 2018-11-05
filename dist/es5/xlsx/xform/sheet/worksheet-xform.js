/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../../../utils/under-dash');

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var RelType = require('../../rel-type');

var Merges = require('./merges');

var BaseXform = require('../base-xform');
var ListXform = require('../list-xform');
var RowXform = require('./row-xform');
var ColXform = require('./col-xform');
var DimensionXform = require('./dimension-xform');
var HyperlinkXform = require('./hyperlink-xform');
var MergeCellXform = require('./merge-cell-xform');
var DataValidationsXform = require('./data-validations-xform');
var SheetPropertiesXform = require('./sheet-properties-xform');
var SheetFormatPropertiesXform = require('./sheet-format-properties-xform');
var SheetViewXform = require('./sheet-view-xform');
var PageMarginsXform = require('./page-margins-xform');
var PageSetupXform = require('./page-setup-xform');
var PrintOptionsXform = require('./print-options-xform');
var AutoFilterXform = require('./auto-filter-xform');
var PictureXform = require('./picture-xform');
var DrawingXform = require('./drawing-xform');
var RowBreaksXform = require('./row-breaks-xform');

var WorkSheetXform = module.exports = function (options) {
  var maxRows = options && options.maxRows;
  this.map = {
    sheetPr: new SheetPropertiesXform(),
    dimension: new DimensionXform(),
    sheetViews: new ListXform({ tag: 'sheetViews', count: false, childXform: new SheetViewXform() }),
    sheetFormatPr: new SheetFormatPropertiesXform(),
    cols: new ListXform({ tag: 'cols', count: false, childXform: new ColXform() }),
    sheetData: new ListXform({ tag: 'sheetData', count: false, empty: true, childXform: new RowXform(), maxItems: maxRows }),
    autoFilter: new AutoFilterXform(),
    mergeCells: new ListXform({ tag: 'mergeCells', count: true, childXform: new MergeCellXform() }),
    rowBreaks: new RowBreaksXform(),
    hyperlinks: new ListXform({ tag: 'hyperlinks', count: false, childXform: new HyperlinkXform() }),
    pageMargins: new PageMarginsXform(),
    dataValidations: new DataValidationsXform(),
    pageSetup: new PageSetupXform(),
    printOptions: new PrintOptionsXform(),
    picture: new PictureXform(),
    drawing: new DrawingXform()
  };
};

utils.inherits(WorkSheetXform, BaseXform, {
  WORKSHEET_ATTRIBUTES: {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x14ac',
    'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
  }
}, {
  prepare: function prepare(model, options) {
    options.merges = new Merges();
    model.hyperlinks = options.hyperlinks = [];

    options.formulae = {};
    options.siFormulae = 0;
    this.map.cols.prepare(model.cols, options);
    this.map.sheetData.prepare(model.rows, options);

    model.mergeCells = options.merges.mergeCells;

    // prepare relationships
    var rels = model.rels = [];
    function nextRid(r) {
      return 'rId' + (r.length + 1);
    }
    var rId;
    model.hyperlinks.forEach(function (hyperlink) {
      rId = nextRid(rels);
      hyperlink.rId = rId;
      rels.push({
        Id: rId,
        Type: RelType.Hyperlink,
        Target: hyperlink.target,
        TargetMode: 'External'
      });
    });

    var drawingRelsHash = [];
    var bookImage;
    model.media.forEach(function (medium) {
      if (medium.type === 'background') {
        rId = nextRid(rels);
        bookImage = options.media[medium.imageId];
        rels.push({
          Id: rId,
          Type: RelType.Image,
          Target: '../media/' + bookImage.name + '.' + bookImage.extension
        });
        model.background = {
          rId: rId
        };
        model.image = options.media[medium.imageId];
      } else if (medium.type === 'image') {
        var drawing = model.drawing;
        bookImage = options.media[medium.imageId];
        if (!drawing) {
          drawing = model.drawing = {
            rId: nextRid(rels),
            name: 'drawing' + ++options.drawingsCount,
            anchors: [],
            rels: []
          };
          options.drawings.push(drawing);
          rels.push({
            Id: drawing.rId,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
            Target: '../drawings/' + drawing.name + '.xml'
          });
        }
        var rIdImage = drawingRelsHash[medium.imageId];
        if (!rIdImage) {
          rIdImage = nextRid(drawing.rels);
          drawingRelsHash[medium.imageId] = rIdImage;
          drawing.rels.push({
            Id: rIdImage,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            Target: '../media/' + bookImage.name + '.' + bookImage.extension
          });
        }
        drawing.anchors.push({
          picture: {
            rId: rIdImage
          },
          range: medium.range
        });
      }
    });
  },

  render: function render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('worksheet', WorkSheetXform.WORKSHEET_ATTRIBUTES);

    var sheetFormatPropertiesModel = model.properties ? {
      defaultRowHeight: model.properties.defaultRowHeight,
      dyDescent: model.properties.dyDescent,
      outlineLevelCol: model.properties.outlineLevelCol,
      outlineLevelRow: model.properties.outlineLevelRow
    } : undefined;
    var sheetPropertiesModel = {
      outlineProperties: model.properties && model.properties.outlineProperties,
      tabColor: model.properties && model.properties.tabColor,
      pageSetup: model.pageSetup && model.pageSetup.fitToPage ? {
        fitToPage: model.pageSetup.fitToPage
      } : undefined
    };
    var pageMarginsModel = model.pageSetup && model.pageSetup.margins;
    var printOptionsModel = {
      showRowColHeaders: model.showRowColHeaders,
      showGridLines: model.showGridLines,
      horizontalCentered: model.horizontalCentered,
      verticalCentered: model.verticalCentered
    };

    this.map.sheetPr.render(xmlStream, sheetPropertiesModel);
    this.map.dimension.render(xmlStream, model.dimensions);
    this.map.sheetViews.render(xmlStream, model.views);
    this.map.sheetFormatPr.render(xmlStream, sheetFormatPropertiesModel);
    this.map.cols.render(xmlStream, model.cols);
    this.map.sheetData.render(xmlStream, model.rows);
    this.map.autoFilter.render(xmlStream, model.autoFilter);
    this.map.mergeCells.render(xmlStream, model.mergeCells);
    this.map.dataValidations.render(xmlStream, model.dataValidations);
    // For some reason hyperlinks have to be after the data validations
    this.map.hyperlinks.render(xmlStream, model.hyperlinks);
    this.map.pageMargins.render(xmlStream, pageMarginsModel);
    this.map.printOptions.render(xmlStream, printOptionsModel);
    this.map.pageSetup.render(xmlStream, model.pageSetup);
    this.map.rowBreaks.render(xmlStream, model.rowBreaks);
    this.map.drawing.render(xmlStream, model.drawing); // Note: must be after rowBreaks
    this.map.picture.render(xmlStream, model.background); // Note: must be after drawing

    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    if (node.name === 'worksheet') {
      _.each(this.map, function (xform) {
        xform.reset();
      });
      return true;
    }

    this.parser = this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen(node);
    }
    return true;
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'worksheet':
        var properties = this.map.sheetFormatPr.model;
        if (this.map.sheetPr.model && this.map.sheetPr.model.tabColor) {
          properties.tabColor = this.map.sheetPr.model.tabColor;
        }
        if (this.map.sheetPr.model && this.map.sheetPr.model.outlineProperties) {
          properties.outlineProperties = this.map.sheetPr.model.outlinePropertiesx;
        }
        var sheetProperties = {
          fitToPage: this.map.sheetPr.model && this.map.sheetPr.model.pageSetup && this.map.sheetPr.model.pageSetup.fitToPage || false,
          margins: this.map.pageMargins.model
        };
        var pageSetup = Object.assign(sheetProperties, this.map.pageSetup.model, this.map.printOptions.model);
        this.model = {
          dimensions: this.map.dimension.model,
          cols: this.map.cols.model,
          rows: this.map.sheetData.model,
          mergeCells: this.map.mergeCells.model,
          hyperlinks: this.map.hyperlinks.model,
          dataValidations: this.map.dataValidations.model,
          properties: properties,
          views: this.map.sheetViews.model,
          pageSetup: pageSetup,
          background: this.map.picture.model,
          drawing: this.map.drawing.model
        };

        if (this.map.autoFilter.model) {
          this.model.autoFilter = this.map.autoFilter.model;
        }
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  },

  reconcile: function reconcile(model, options) {
    // options.merges = new Merges();
    // options.merges.reconcile(model.mergeCells, model.rows);
    var rels = (model.relationships || []).reduce(function (h, rel) {
      h[rel.Id] = rel;
      return h;
    }, {});
    options.hyperlinkMap = (model.hyperlinks || []).reduce(function (h, hyperlink) {
      if (hyperlink.rId) {
        h[hyperlink.address] = rels[hyperlink.rId].Target;
      }
      return h;
    }, {});
    options.formulae = {};

    // compact the rows and cells
    model.rows = model.rows && model.rows.filter(Boolean) || [];
    model.rows.forEach(function (row) {
      row.cells = row.cells && row.cells.filter(Boolean) || [];
    });

    this.map.cols.reconcile(model.cols, options);
    this.map.sheetData.reconcile(model.rows, options);

    model.media = [];
    if (model.drawing) {
      var drawingRel = rels[model.drawing.rId];
      var match = drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
      if (match) {
        var drawingName = match[1];
        var drawing = options.drawings[drawingName];
        drawing.anchors.forEach(function (anchor) {
          if (anchor.medium) {
            var image = {
              type: 'image',
              imageId: anchor.medium.index,
              range: anchor.range
            };
            model.media.push(image);
          }
        });
      }
    }

    var backgroundRel = model.background && rels[model.background.rId];
    if (backgroundRel) {
      var target = backgroundRel.Target.split('/media/')[1];
      var imageId = options.mediaIndex && options.mediaIndex[target];
      if (imageId !== undefined) {
        model.media.push({
          type: 'background',
          imageId: imageId
        });
      }
    }

    delete model.relationships;
    delete model.hyperlinks;
  }
});
//# sourceMappingURL=worksheet-xform.js.map

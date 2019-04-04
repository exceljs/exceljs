'use strict';

const utils = require('../../../utils/utils');
const XmlStream = require('../../../utils/xml-stream');

const BaseXform = require('../base-xform');
const TwoCellAnchorXform = require('./two-cell-anchor-xform');

const WorkSheetXform = (module.exports = function() {
  this.map = {
    'xdr:twoCellAnchor': new TwoCellAnchorXform(),
  };
});

utils.inherits(
  WorkSheetXform,
  BaseXform,
  {
    DRAWING_ATTRIBUTES: {
      'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    },
  },
  {
    get tag() {
      return 'xdr:wsDr';
    },

    prepare(model) {
      const twoCellAnchorXform = this.map['xdr:twoCellAnchor'];
      model.anchors.forEach((item, index) => {
        twoCellAnchorXform.prepare(item, { index });
      });
    },

    render(xmlStream, model) {
      xmlStream.openXml(XmlStream.StdDocAttributes);
      xmlStream.openNode(this.tag, WorkSheetXform.DRAWING_ATTRIBUTES);

      const twoCellAnchorXform = this.map['xdr:twoCellAnchor'];
      model.anchors.forEach(item => {
        twoCellAnchorXform.render(xmlStream, item);
      });

      xmlStream.closeNode();
    },

    parseOpen(node) {
      if (this.parser) {
        this.parser.parseOpen(node);
        return true;
      }
      switch (node.name) {
        case this.tag:
          this.reset();
          this.model = {
            anchors: [],
          };
          break;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
          break;
      }
      return true;
    },

    parseText(text) {
      if (this.parser) {
        this.parser.parseText(text);
      }
    },

    parseClose(name) {
      if (this.parser) {
        if (!this.parser.parseClose(name)) {
          this.model.anchors.push(this.parser.model);
          this.parser = undefined;
        }
        return true;
      }
      switch (name) {
        case this.tag:
          return false;
        default:
          // could be some unrecognised tags
          return true;
      }
    },

    reconcile(model, options) {
      model.anchors.forEach(anchor => {
        this.map['xdr:twoCellAnchor'].reconcile(anchor, options);
      });
    },
  }
);

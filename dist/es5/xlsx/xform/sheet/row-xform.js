/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var CellXform = require('./cell-xform');

var RowXform = module.exports = function () {
  this.map = {
    c: new CellXform()
  };
};

// <row r="<%=row.number%>"
//     <% if(row.height) {%> ht="<%=row.height%>" customHeight="1"<% } %>
//     <% if(row.hidden) {%> hidden="1"<% } %>
//     <% if(row.min > 0 && row.max > 0 && row.min <= row.max) {%> spans="<%=row.min%>:<%=row.max%>"<% } %>
//     <% if (row.styleId) { %> s="<%=row.styleId%>" customFormat="1" <% } %>
//     x14ac:dyDescent="0.25">
//   <% row.cells.forEach(function(cell){ %>
//     <%=cell.xml%><% }); %>
// </row>

utils.inherits(RowXform, BaseXform, {

  get tag() {
    return 'row';
  },

  prepare: function prepare(model, options) {
    var styleId = options.styles.addStyleModel(model.style);
    if (styleId) {
      model.styleId = styleId;
    }
    var cellXform = this.map.c;
    model.cells.forEach(function (cellModel) {
      cellXform.prepare(cellModel, options);
    });
  },

  render: function render(xmlStream, model, options) {
    xmlStream.openNode('row');
    xmlStream.addAttribute('r', model.number);
    if (model.height) {
      xmlStream.addAttribute('ht', model.height);
      xmlStream.addAttribute('customHeight', '1');
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden', '1');
    }
    if (model.min > 0 && model.max > 0 && model.min <= model.max) {
      xmlStream.addAttribute('spans', model.min + ':' + model.max);
    }
    if (model.styleId) {
      xmlStream.addAttribute('s', model.styleId);
      xmlStream.addAttribute('customFormat', '1');
    }
    xmlStream.addAttribute('x14ac:dyDescent', '0.25');
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel', model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed', '1');
    }

    var cellXform = this.map.c;
    model.cells.forEach(function (cellModel) {
      cellXform.render(xmlStream, cellModel, options);
    });

    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (node.name === 'row') {
      this.numRowsSeen += 1;
      var spans = node.attributes.spans ? node.attributes.spans.split(':').map(function (span) {
        return parseInt(span, 10);
      }) : [undefined, undefined];
      var model = this.model = {
        number: parseInt(node.attributes.r, 10),
        min: spans[0],
        max: spans[1],
        cells: []
      };
      if (node.attributes.s) {
        model.styleId = parseInt(node.attributes.s, 10);
      }
      if (node.attributes.hidden) {
        model.hidden = true;
      }
      if (node.attributes.bestFit) {
        model.bestFit = true;
      }
      if (node.attributes.ht) {
        model.height = parseFloat(node.attributes.ht);
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel, 10);
      }
      if (node.attributes.collapsed) {
        model.collapsed = true;
      }
      return true;
    }

    this.parser = this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.cells.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    return false;
  },

  reconcile: function reconcile(model, options) {
    model.style = model.styleId ? options.styles.getStyleModel(model.styleId) : {};
    if (model.styleId !== undefined) {
      model.styleId = undefined;
    }

    var cellXform = this.map.c;
    model.cells.forEach(function (cellModel) {
      cellXform.reconcile(cellModel, options);
    });
  }
});
//# sourceMappingURL=row-xform.js.map

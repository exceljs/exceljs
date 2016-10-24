/**
 * Copyright (c) 2015 Guyon Roche
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

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var CellXform = require('./cell-xform');

var RowXform = module.exports = function() {
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

  get tag() { return 'row'; },

  prepare: function(model, options) {
    var styleId = options.styles.addStyleModel(model.style);
    if (styleId) {
      model.styleId = styleId;
    }
    var cellXform = this.map.c;
    model.cells.forEach(function(cellModel) {
      cellXform.prepare(cellModel, options);
    });
  },
  
  render: function(xmlStream, model, options) {
    xmlStream.openNode('row');
    xmlStream.addAttribute('r',  model.number);
    if (model.height) {
      xmlStream.addAttribute('ht',  model.height);
      xmlStream.addAttribute('customHeight', '1');
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden',  '1');
    }
    if (model.min > 0 && model.max > 0 && model.min <= model.max) {
      xmlStream.addAttribute('spans', model.min + ':' + model.max);
    }
    if (model.styleId) {
      xmlStream.addAttribute('s',  model.styleId);
      xmlStream.addAttribute('customFormat', '1');
    }
    xmlStream.addAttribute('x14ac:dyDescent', '0.25');
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel',  model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed',  '1');
    }

    var cellXform = this.map.c;
    model.cells.forEach(function(cellModel) {
      cellXform.render(xmlStream, cellModel, options);
    });

    xmlStream.closeNode();
  },
  
  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (node.name === 'row') {
      var spans = node.attributes.spans ? node.attributes.spans.split(':').map(function(span) { return parseInt(span); }) : [undefined, undefined];
      var model = this.model = {
        number: parseInt(node.attributes.r),
        min: spans[0],
        max: spans[1],
        cells: []
      };
      if (node.attributes.s) {
        model.styleId = parseInt(node.attributes.s);
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
        model.outlineLevel = parseInt(node.attributes.outlineLevel);
      }
      if (node.attributes.collapsed) {
        model.collapsed = true;
      }
      return true;
    } else if ((this.parser = this.map[node.name])) {
      this.parser.parseOpen(node);
      return true;
    } else {
      return false;
    }
  },
  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.cells.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    } else {
      return false;
    }
  },
  
  reconcile: function(model, options) {
    model.style = model.styleId ? options.styles.getStyleModel(model.styleId) : {};
    if (model.styleId != undefined) {
      model.styleId = undefined;
    }

    var cellXform = this.map.c;
    model.cells.forEach(function(cellModel) {
      cellXform.reconcile(cellModel, options);
    });
  }
});

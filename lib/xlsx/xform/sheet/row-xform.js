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
var BaseXform = require('./../base-xform');

var CellXform = require('./cell-xform');

var RowXform = module.exports = function() {
  this.map = {
    c: new CellXform()
  };
};

// <row r="<%=row.number%>"
//     <% if(row.height) {%> ht="<%=row.height%>" customHeight="1"<% } %>
//     <% if(row.hidden) {%> hidden="1"<% } %>
//     spans="<%=row.min%>:<%=row.max%>"
//     <% if (row.styleId) { %> s="<%=row.styleId%>" customFormat="1" <% } %>
//     x14ac:dyDescent="0.25">
//   <% _.each(row.cells, function(cell){ %>
//     <%=cell.xml%><% }); %>
// </row>

utils.inherits(RowXform, BaseXform, {
  
  prepare: function(model, options) {
    var cellXform = this.map.c;
    model.cells.forEach(function(cellModel) {
      cellXform.prepare(cellModel, options);
    });
  },
  
  write: function(xmlStream, model, options) {
    xmlStream.openNode('row');
    xmlStream.addAttribute('r',  model.number);
    if (model.height) {
      xmlStream.addAttribute('height',  model.height);
      xmlStream.addAttribute('customHeight', '1');
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden',  '1');
    }
    xmlStream.addAttribute('spans', model.min + ':' + model.max);
    if (model.styleId) {
      xmlStream.addAttribute('s',  model.styleId);
      xmlStream.addAttribute('customFormat', '1');
    }
    xmlStream.addAttribute('x14ac:dyDescent', '0.25');

    var cellXform = this.map.c;
    model.cells.forEach(function(cellModel) {
      cellXform.write(xmlStream, cellModel, options);
    });

    xmlStream.closeNode();
  },
  
  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (node.name === 'row') {
      var model = this.model = {
        min: parseInt(node.attributes.min || '0'),
        max: parseInt(node.attributes.max || '0'),
        width: parseFloat(node.attributes.width || '0'),
        cells: []
      };
      if (node.attributes.style) {
        model.styleId = parseInt(node.attributes.style);
      }
      if (node.attributes.hidden) {
        model.hidden = true;
      }
      if (node.attributes.bestFit) {
        model.bestFit = true;
      }
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
    var cellXform = this.map.c;
    model.cells.forEach(function(cellModel) {
      cellXform.reconcile(cellModel, options);
    });
  }
});

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

var ColorXform = require('./color-xform');

//var model = {
//    color: { argb: "12345678" }, // default colour for individual directions
//    top: { style: "thick" }, // gets default colour
//    left: { style: "thin", color:{auto:true} }, // overrides with auto colour
//    diagonal: { up: true, down: true, style: "thin" }
//}

var EdgeXform = function(name) {
  this.name = name;
  this.map = {
    color: new ColorXform()
  };
};

utils.inherits(EdgeXform, BaseXform, {

  get tag() { return this.name; },

  render: function(xmlStream, model, defaultColor) {
    var color = model && model.color || defaultColor || this.defaultColor;
    xmlStream.openNode(this.name);
    if (model && model.style) {
      xmlStream.addAttribute('style', model.style);
      if (color) {
        this.map.color.render(xmlStream, color);
      }
    }
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch (node.name) {
        case this.name:
          var style = node.attributes.style;
          if (style) {
            this.model = {
              style: style
            };
          } else {
            this.model = undefined;
          }
          return true;
        case 'color':
          this.parser  = this.map.color;
          this.parser.parseOpen(node);
          return true;
        default:
          return false;
      }
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
        this.parser = undefined;
      }
      return true;
    } else {
      if (name === this.name) {
        if (this.map.color.model) {
          if(!this.model)
            this.model={};
          this.model.color = this.map.color.model;
        }
      }
      return false;
    }
  },

  validStyleValues: [
    'thin',
    'dotted',
    'dashDot',
    'hair',
    'dashDotDot',
    'slantDashDot',
    'mediumDashed',
    'mediumDashDotDot',
    'mediumDashDot',
    'medium',
    'double',
    'thick'
  ].reduce(function(p,v){p[v]=true; return p;}, {}),
  validStyle: function(value) {
    return this.validStyleValues[value];
  }
});

// Border encapsulates translation from border model to/from xlsx
var BorderXform = module.exports = function() {
  this.map = {
    top: new EdgeXform('top'),
    left: new EdgeXform('left'),
    bottom: new EdgeXform('bottom'),
    right: new EdgeXform('right'),
    diagonal: new EdgeXform('diagonal')
  };
};

utils.inherits(BorderXform, BaseXform, {
  render: function(xmlStream, model) {
    var color = model.color;
    xmlStream.openNode('border');
    if (model.diagonal && model.diagonal.style) {
      if (model.diagonal.up) {
        xmlStream.addAttribute('diagonalUp', '1');
      }
      if (model.diagonal.down) {
        xmlStream.addAttribute('diagonalDown', '1');
      }
    }
    function add(edgeModel, edgeXform) {
      if (edgeModel && !edgeModel.color && model.color) {
        // don't mess with incoming models
        edgeModel = Object.assign({}, edgeModel, {color: model.color});
      }
      edgeXform.render(xmlStream, edgeModel, color);
    }
    add(model.left, this.map.left);
    add(model.right, this.map.right);
    add(model.top, this.map.top);
    add(model.bottom, this.map.bottom);
    add(model.diagonal, this.map.diagonal);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
    } else {
      switch(node.name) {
        case 'border':
          this.diagonalUp = node.attributes.diagonalUp ? true : false;
          this.diagonalDown = node.attributes.diagonalDown ? true : false;
          this.map.left.reset();
          this.map.right.reset();
          this.map.top.reset();
          this.map.bottom.reset();
          this.map.diagonal.reset();
          return true;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
            return true;
          } else {
            return false;
          }
      }
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
        this.parser = undefined;
      }
      return true;
    } else {
      if (name === 'border') {
        var model = this.model = {};
        var add = function(name, edgeModel, extensions) {
          if (edgeModel) {
            if (extensions) {
              Object.assign(edgeModel, extensions);
            }
            model[name] = edgeModel;
          }
        };
        add('left', this.map.left.model);
        add('right', this.map.right.model);
        add('top', this.map.top.model);
        add('bottom', this.map.bottom.model);
        add('diagonal', this.map.diagonal.model, {up: this.diagonalUp, down: this.diagonalDown});
      }
      return false;
    }
  }
});

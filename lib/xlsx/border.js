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

var _ = require('underscore');

var ColorXform = require('./xform/color-xform');

//var model = {
//    color: { argb: "12345678" }, // default colour for individual directions
//    top: { style: "thick" }, // gets default colour
//    left: { style: "thin", color:{auto:true} }, // overrides with auto colour
//
//}

var Edge = function(name, isDiagonal, model, defaultColor) {
  this.name = name;
  this.isDiagonal = isDiagonal;
  this.color = defaultColor;
  if (model) {
    this.style = model.style;
    if (this.style) {
      if (model.color) {
        this.color = new ColorXform(model.color);
      }
      if (isDiagonal) {
        this.up = model.up;
        this.down = model.down;
        if (!(this.up || this.down)) {
          this.style = null;
        }
      }
    }
  }
}
Edge.prototype = {
  get model() {
    if (this.style) {
      var model = {
        style: this.style
      };
      if (this.color) {
        var colorModel = this.color.model;
        if (colorModel) { model.color = colorModel; }
      }
      if (this.isDiagonal) {
        model.up = this.up ? true : false;
        model.down = this.down ? true : false;
      }
      return model;
    } else {
      return null;
    }
  },
  get xml() {
    //<right style="thick">
    //    <color theme="9" tint="0.59996337778862885"/>
    //</right>
    if (this.style) {
      return [
        '<', this.name, ' style="', this.style, '">',
        (this.color || new ColorXform()).xml,
        '</', this.name, '>'
      ].join('');
    } else {
      return '<' + this.name + '/>';
    }
  },
  parse: function(node) {
    switch(node.name) {
      case this.name:
        this.style = node.attributes.style;
        break;
      case 'color':
        this.color = new ColorXform();
        this.color.parseOpen(node);
        break;
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
};

// Border encapsulates translation from border model to/from xlsx
var Border = module.exports = function(model) {
  if (model) {
    this.color = new ColorXform(model.color);
    this.top = new Edge('top', false, model.top, this.color);
    this.left = new Edge('left', false, model.left, this.color);
    this.bottom = new Edge('bottom', false, model.bottom, this.color);
    this.right = new Edge('right', false, model.right, this.color);
    this.diagonal = new Edge('diagonal', true, model.diagonal, this.color);
  }
};

Border.prototype = {
  get model() {
    var model = {};
    var hasValues = false;
    function add(name, value) {
      if (value && value.model) {
        model[name] = value.model;
        hasValues = true;
      }
    }
    add('color', this.color);
    add('top', this.top);
    add('left', this.left);
    add('bottom', this.bottom);
    add('right', this.right);
    add('diagonal', this.diagonal);

    return hasValues ? model : null;
  },

  get xml() {
    // return string containing the <border> definition
    if (this._xml === undefined) {
      if (this.left) {
        var xml = ['<border'];
        // left, right, top bottom
        if (this.diagonal.style) {
          if (this.diagonal.up) {
            xml.push(' diagonalUp="1"');
          }
          if (this.diagonal.down) {
            xml.push(' diagonalDown="1"');
          }
        }
        xml.push('>');
        xml.push(this.left.xml);
        xml.push(this.right.xml);
        xml.push(this.top.xml);
        xml.push(this.bottom.xml);
        xml.push(this.diagonal.xml);
        xml.push('</border>');
        this._xml = xml.join('');
      } else {
        // special case - empty border
        this._xml = '<border><left/><right/><top/><bottom/><diagonal/></border>';
      }
    }
    return this._xml;
  },

  parse: function(node) {
    // used during sax parsing of xml to build border object
    switch(node.name) {
      case 'border':
        if (node.attributes.diagonalUp || node.attributes.diagonalDown) {
          this.diagonal = new Edge('diagonal', true);
          this.diagonal.up = node.attributes.diagonalUp ? true : false;
          this.diagonal.down = node.attributes.diagonalDown ? true : false;
        }
        break;
      case 'left':
        this.parsing = this.left = new Edge('left', false);
        break;
      case 'right':
        this.parsing = this.right = new Edge('right', false);
        break;
      case 'top':
        this.parsing = this.top = new Edge('top', false);
        break;
      case 'bottom':
        this.parsing = this.bottom = new Edge('bottom', false);
        break;
      case 'diagonal':
        if (!this.diagonal) {
          this.diagonal = new Edge('diagonal', true);
        }
        this.parsing = this.diagonal;
        break;
    }
    // if node is <color>, pass to last created edge
    if (this.parsing) {
      this.parsing.parse(node);
    }
  }
};

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

// pattern fill: fgColor is pattern, gbColor is solid gb fill
//<fill>
//    <patternFill patternType="solid">
//        <fgColor theme="0" tint="-4.9989318521683403E-2"/>
//        <bgColor indexed="64"/>
//    </patternFill>
//</fill>
//
// angle gradient. degree=[0-180] 0=L->R, 1-359=rotate clockwise
// stop sequence specifies way points for the color to travel through
//<fill>
//    <gradientFill degree="30">
//        <stop position="0"><color theme="0"/></stop>
//        <stop position="1"><color theme="4"/></stop>
//    </gradientFill>
//</fill>
//
// "path" gradient. Outwards from a point (or rectangle) where left=right=x, top=bottom=y
//<fill>
//    <gradientFill type="path" left="0.3" right="0.3" top="0.7" bottom="0.7">
//        <stop position="0"><color rgb="FFFF0000"/></stop>
//        <stop position="0.5"><color rgb="FF00FF00"/></stop>
//        <stop position="1"><color rgb="FF0000FF"/></stop>
//    </gradientFill>
//</fill>

var modelSchema = {
  // fill will be one of these types
  type: ['pattern', 'gradient'],

  // ============================================================
  // pattern properties

  // pattern style, required
  pattern: ['solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625',
    'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
    'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'],

  // foreground colour, default is auto (i.e. black)
  fgColor: new ColorXform('fgColor'),

  // background (solid fill) colour, default is auto (i.e. white)
  bgColor: new ColorXform('bgColor'),

  // ============================================================
  // gradient properties
  gradient: ['angle', 'path'],
  degree: [0,359], // for angle gradients 0=left to right, then clockwise from that
  center: { left:[0,1], top:[0,1] }, // for path gradients optional right, bottom for a square
  stops: [ //  start at position 0 and end at position 1
    {position:0, color: {argb:''}},
    {position:(0,1), color: {argb:''}}, // intermediate steps optional
    {position:1, color: {argb:''}}
  ]
};

var Stop = function(model) {
  if (model) {
    this.position = model.position;
    this.color = new ColorXform(model.color);
  }
};
Stop.prototype = {
  get model() {
    return {
      position: this.position,
      color: this.color.model
    };
  },
  get xml() {
    return '<stop position="' + this.position + '">' + this.color.xml + '</stop>';
  },
  parse: function(node) {
    switch(node.name) {
      case 'stop':
        this.position = parseFloat(node.attributes.position);
        break;
      case 'color':
        this.color = new ColorXform();
        this.color.parseOpen(node);
        break;
    }
  }
};

var PatternFill = function(model) {
  if (model) {
    this.pattern = model.pattern;
    if (model.fgColor) {
      this.fgColor = new ColorXform(model.fgColor, 'fgColor');
    }
    if (model.bgColor) {
      this.bgColor = new ColorXform(model.bgColor, 'bgColor');
    }
  }
};
PatternFill.prototype = {
  get model() {
    var model = {
      type: 'pattern',
      pattern: this.pattern
    };
    if (this.fgColor) {
      model.fgColor = this.fgColor.model;
    }
    if (this.bgColor) {
      model.bgColor = this.bgColor.model;
    }
    return model;
  },
  get xml() {
    if (this._xml === undefined) {
      var xml = [];
      xml.push('<patternFill patternType="');
      xml.push(this.pattern);
      xml.push('"');
      if (this.fgColor || this.bgColor) {
        xml.push('>');
        if (this.fgColor) {
          xml.push(this.fgColor.xml);
        }
        if (this.bgColor) {
          xml.push(this.bgColor.xml);
        }
        xml.push('</patternFill>');
      } else {
        xml.push('/>');
      }
      this._xml = xml.join('');
    }
    return this._xml;
  },
  parse: function(node) {
    switch(node.name) {
      case 'patternFill':
        this.pattern = node.attributes.patternType;
        break;
      case 'fgColor':
        this.fgColor = new ColorXform(null, 'fgColor');
        this.fgColor.parseOpen(node);
        break;
      case 'bgColor':
        this.bgColor = new ColorXform(null, 'bgColor');
        this.bgColor.parseOpen(node);
        break;
    }
  }
};

var GradientFill = function(model) {
  if (model) {
    this.gradient = model.gradient;
    if (model.center) {
      this.center = model.center;
    }
    if (model.degree !== undefined) {
      this.degree = model.degree;
    }
    this.stops = model.stops.map(function(stop) { return new Stop(stop); });
  } else {
    this.stops = [];
  }
}
GradientFill.prototype = {
  get model() {
    switch(this.gradient) {
      case 'angle':
        return {
          type: 'gradient',
          gradient: this.gradient,
          degree: this.degree,
          stops: this.stops.map(function(stop) { return stop.model; })
        };
      case 'path':
        return {
          type: 'gradient',
          gradient: this.gradient,
          center: this.center,
          stops: this.stops.map(function(stop) { return stop.model; })
        };
      default:
        return null;
    }
  },

  get xml() {
    if (this._xml === undefined) {
      var xml = [];
      xml.push('<gradientFill');

      switch(this.gradient) {
        case 'angle':
          xml.push(' degree="' + this.degree + '">');
          break;
        case 'path':
          xml.push(' type="path"');
          if (this.center.left) {
            xml.push(' left="' + this.center.left + '"');
            if (this.center.right === undefined) {
              xml.push(' right="' + this.center.left + '"');
            }
          }
          if (this.center.right) {
            xml.push(' right="' + this.center.right + '"');
          }
          if (this.center.top) {
            xml.push(' top="' + this.center.top + '"');
            if (this.center.bottom === undefined) {
              xml.push(' bottom="' + this.center.top + '"');
            }
          }
          if (this.center.bottom) {
            xml.push(' bottom="' + this.center.bottom + '"');
          }
          xml.push('>');
          break;
      }

      _.each(this.stops, function(stop) {
        xml.push(stop.xml);
      });
      xml.push('</gradientFill>');
      this._xml = xml.join('');
    }
    return this._xml;
  },
  parse: function(node) {
    switch(node.name) {
      case 'gradientFill':
        if (node.attributes.degree) {
          this.gradient = 'angle';
          this.degree = parseInt(node.attributes.degree);
        } else if (node.attributes.type == 'path') {
          this.gradient = 'path';
          this.center = {
            left: node.attributes.left ? parseFloat(node.attributes.left) : 0,
            top: node.attributes.top ? parseFloat(node.attributes.top) : 0
          }
          if (node.attributes.right != node.attributes.left) {
            this.center.right = node.attributes.right ? parseFloat(node.attributes.right) : 0;
          }
          if (node.attributes.bottom != node.attributes.top) {
            this.center.bottom = node.attributes.bottom ? parseFloat(node.attributes.bottom) : 0;
          }
        }
        break;
      case 'stop':
        this.stop = new Stop();
        this.stops.push(this.stop);
        this.stop.parse(node);
        break;
      default:
        // internals of stop
        if (this.stop) {
          this.stop.parse(node);
        }
    }
  }
};


// Fill encapsulates translation from fill model to/from xlsx
var Fill = module.exports = function(model) {
  if (model) {
    if (model.pattern) {
      this._model = new PatternFill(model);
    } else if (model.gradient) {
      this._model = new GradientFill(model);
    } else {
      throw new Error('Did not recognise this fill model: ' + JSON.stringify(model));
    }
  }
};

Fill.prototype = {
  get model() {
    return this._model ? this._model.model : null;
  },

  get xml() {
    if (this._xml === undefined) {
      var innerXml = this._model && this._model.xml;
      if (innerXml) {
        this._xml = '<fill>' + innerXml + '</fill>';
      } else {
        this._xml = null;
      }
    }
    return this._xml;
  },

  parse: function(node) {
    // used during sax parsing of xml to build fill object
    switch(node.name) {
      case 'patternFill':
        this._model = new PatternFill();
        break;
      case 'gradientFill':
        this._model = new GradientFill();
        break;
    }
    if (this._model) {
      this._model.parse(node);
    }
  },

  validPatternValues: [
    'none',
    'solid',
    'darkVertical',
    'darkGray',
    'mediumGray',
    'lightGray',
    'gray125',
    'gray0625',
    'darkHorizontal',
    'darkVertical',
    'darkDown',
    'darkUp',
    'darkGrid',
    'darkTrellis',
    'lightHorizontal',
    'lightVertical',
    'lightDown',
    'lightUp',
    'lightGrid',
    'lightTrellis',
    'lightGrid'
  ].reduce(function(p,v){p[v]=true; return p;}, {}),
  validStyle: function(value) {
    return this.validStyleValues[value];
  }
};

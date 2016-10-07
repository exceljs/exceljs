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

// var modelSchema = {
//   // fill will be one of these types
//   type: ['pattern', 'gradient'],
//
//   // ============================================================
//   // pattern properties
//
//   // pattern style, required
//   pattern: ['solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625',
//     'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
//     'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'],
//
//   // foreground colour, default is auto (i.e. black)
//   fgColor: new ColorXform('fgColor'),
//
//   // background (solid fill) colour, default is auto (i.e. white)
//   bgColor: new ColorXform('bgColor'),
//
//   // ============================================================
//   // gradient properties
//   gradient: ['angle', 'path'],
//   degree: [0,359], // for angle gradients 0=left to right, then clockwise from that
//   center: { left:[0,1], top:[0,1] }, // for path gradients optional right, bottom for a square
//   stops: [ //  start at position 0 and end at position 1
//     {position:0, color: {argb:''}},
//     {position:(0,1), color: {argb:''}}, // intermediate steps optional
//     {position:1, color: {argb:''}}
//   ]
// };

var StopXform = function() {
  this.map = {
    color: new ColorXform()
  };
};

utils.inherits(StopXform, BaseXform, {

  get tag() { return 'stop'; },

  render:  function(xmlStream, model) {
    xmlStream.openNode('stop');
    xmlStream.addAttribute('position', model.position);
    this.map.color.render(xmlStream, model.color);
    xmlStream.closeNode();
  },
  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch (node.name) {
        case 'stop':
          this.model = {
            position: parseFloat(node.attributes.position)
          };
          return true;
        case 'color':
          this.parser = this.map.color;
          this.parser.parseOpen(node);
          return true;
        default:
          return false;
      }
    }
  },
  parseText: function() {
  },
  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.color = this.parser.model;
        this.parser = undefined;
      }
      return true;
    }
    return false;
  }
});

var PatternFillXform = function() {
  this.map = {
    fgColor: new ColorXform('fgColor'),
    bgColor: new ColorXform('bgColor')
  };
};
utils.inherits(PatternFillXform, BaseXform, {

  get name() { return 'pattern'; },
  get tag() { return 'patternFill'; },

  render: function(xmlStream, model) {
    xmlStream.openNode('patternFill');
    xmlStream.addAttribute('patternType', model.pattern);
    if (model.fgColor) {
      this.map.fgColor.render(xmlStream, model.fgColor);
    }
    if (model.bgColor) {
      this.map.bgColor.render(xmlStream, model.bgColor);
    }
    xmlStream.closeNode();
  },
  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'patternFill':
          this.model = {
              type: 'pattern',
              pattern: node.attributes.patternType
            };
          return true;
        default:
          if ((this.parser = this.map[node.name])) {
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
        if (this.parser.model) {
          this.model[name] = this.parser.model;
        }
        this.parser = undefined;
      }
      return true;
    } else {
      return false;
    }
  }
});

var GradientFillXform = function() {
  this.map = {
    stop: new StopXform()
  };
  // if (model) {
  //   this.gradient = model.gradient;
  //   if (model.center) {
  //     this.center = model.center;
  //   }
  //   if (model.degree !== undefined) {
  //     this.degree = model.degree;
  //   }
  //   this.stops = model.stops.map(function(stop) { return new StopXform(stop); });
  // } else {
  //   this.stops = [];
  // }
};
utils.inherits(GradientFillXform, BaseXform, {
  get name() { return 'gradient'; },
  get tag() { return 'gradientFill'; },
  
  render: function(xmlStream, model) {
    xmlStream.openNode('gradientFill');
    switch(model.gradient) {
      case 'angle':
        xmlStream.addAttribute('degree', model.degree);
        break;
      case 'path':
        xmlStream.addAttribute('type', 'path');
        if (model.center.left) {
          xmlStream.addAttribute('left', model.center.left);
          if (model.center.right === undefined) {
            xmlStream.addAttribute('right', model.center.left);
          }
        }
        if (model.center.right) {
          xmlStream.addAttribute('right', model.center.right);
        }
        if (model.center.top) {
          xmlStream.addAttribute('top', model.center.top);
          if (model.center.bottom === undefined) {
            xmlStream.addAttribute('bottom', model.center.top);
          }
        }
        if (model.center.bottom) {
          xmlStream.addAttribute('bottom', model.center.bottom);
        }
        break;
    }

    var stopXform = this.map.stop;
    model.stops.forEach(function(stopModel) {
      stopXform.render(xmlStream, stopModel);
    });

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch (node.name) {
        case 'gradientFill':
          var model = this.model = {
            stops: []
          };
          if (node.attributes.degree) {
            model.gradient = 'angle';
            model.degree = parseInt(node.attributes.degree);
          } else if (node.attributes.type == 'path') {
            model.gradient = 'path';
            model.center = {
              left: node.attributes.left ? parseFloat(node.attributes.left) : 0,
              top: node.attributes.top ? parseFloat(node.attributes.top) : 0
            };
            if (node.attributes.right != node.attributes.left) {
              model.center.right = node.attributes.right ? parseFloat(node.attributes.right) : 0;
            }
            if (node.attributes.bottom != node.attributes.top) {
              model.center.bottom = node.attributes.bottom ? parseFloat(node.attributes.bottom) : 0;
            }
          }
          return true;
        case 'stop':
          this.parser = this.map.stop;
          this.parser.parseOpen(node);
          return true;
          break;
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
  parseClose:  function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.stops.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    } else {
      return false;
    }
  }
});


// Fill encapsulates translation from fill model to/from xlsx
var FillXform = module.exports = function() {
  this.map = {
    patternFill:  new PatternFillXform(),
    gradientFill: new GradientFillXform()
  };
  // if (model) {
  //   if (model.pattern) {
  //     this._model = new PatternFillXform(model);
  //   } else if (model.gradient) {
  //     this._model = new GradientFill(model);
  //   } else {
  //     throw new Error('Did not recognise this fill model: ' + JSON.stringify(model));
  //   }
  // }
};

utils.inherits(FillXform, BaseXform, {
  StopXform: StopXform,
  PatternFillXform: PatternFillXform,
  GradientFillXform: GradientFillXform
},{
  get tag() { return 'fill'; },

  render: function(xmlStream, model) {
    xmlStream.addRollback();
    xmlStream.openNode('fill');
    switch(model.type) {
      case 'pattern':
        this.map.patternFill.render(xmlStream, model);
        break;
      case 'gradient':
        this.map.gradientFill.render(xmlStream, model);
        break;
      default:
        xmlStream.rollback();
        return;
    }
    xmlStream.closeNode();
    xmlStream.commit();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch (node.name) {
        case 'fill':
          this.model = {};
          return true;
        default:
          if ((this.parser = this.map[node.name])) {
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
  parseClose:  function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model = this.parser.model;
        this.model.type = this.parser.name;
        this.parser = undefined;
      }
      return true;
    } else {
      return false;
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
});

/**
 * Copyright (c) 2015 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var ColorXform = require('./color-xform');

var EdgeXform = function EdgeXform(name) {
  this.name = name;
  this.map = {
    color: new ColorXform()
  };
};

utils.inherits(EdgeXform, BaseXform, {
  get tag() {
    return this.name;
  },

  render: function render(xmlStream, model, defaultColor) {
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

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
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
        this.parser = this.map.color;
        this.parser.parseOpen(node);
        return true;
      default:
        return false;
    }
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

    if (name === this.name) {
      if (this.map.color.model) {
        if (!this.model) {
          this.model = {};
        }
        this.model.color = this.map.color.model;
      }
    }

    return false;
  },

  validStyleValues: ['thin', 'dotted', 'dashDot', 'hair', 'dashDotDot', 'slantDashDot', 'mediumDashed', 'mediumDashDotDot', 'mediumDashDot', 'medium', 'double', 'thick'].reduce(function (p, v) {
    p[v] = true;return p;
  }, {}),
  validStyle: function validStyle(value) {
    return this.validStyleValues[value];
  }
});

// Border encapsulates translation from border model to/from xlsx
var BorderXform = module.exports = function () {
  this.map = {
    top: new EdgeXform('top'),
    left: new EdgeXform('left'),
    bottom: new EdgeXform('bottom'),
    right: new EdgeXform('right'),
    diagonal: new EdgeXform('diagonal')
  };
};

utils.inherits(BorderXform, BaseXform, {
  render: function render(xmlStream, model) {
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
        edgeModel = Object.assign({}, edgeModel, { color: model.color });
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

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'border':
        this.reset();
        this.diagonalUp = !!node.attributes.diagonalUp;
        this.diagonalDown = !!node.attributes.diagonalDown;
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
        return false;
    }
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
    if (name === 'border') {
      var model = this.model = {};
      var add = function add(key, edgeModel, extensions) {
        if (edgeModel) {
          if (extensions) {
            Object.assign(edgeModel, extensions);
          }
          model[key] = edgeModel;
        }
      };
      add('left', this.map.left.model);
      add('right', this.map.right.model);
      add('top', this.map.top.model);
      add('bottom', this.map.bottom.model);
      add('diagonal', this.map.diagonal.model, { up: this.diagonalUp, down: this.diagonalDown });
    }
    return false;
  }
});
//# sourceMappingURL=border-xform.js.map

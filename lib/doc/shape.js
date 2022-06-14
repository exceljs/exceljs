const colCache = require('../utils/col-cache');
const Anchor = require('./anchor');

class Shape {
  constructor(worksheet, model) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model() {
    switch (this.type) {
      case 'shape':
        return {
          type: this.type,
          shape: this.shape,
          rotation: this.rotation,
          fill: this.fill,
          stroke: this.stroke,
          hyperlinks: this.range.hyperlinks,
          range: {
            tl: this.range.tl.model,
            br: this.range.br && this.range.br.model,
            ext: this.range.ext,
            editAs: this.range.editAs,
          },
        };
      default:
        throw new Error('Invalid Shape Type');
    }
  }

  set model({type, shape, rotation, fill, stroke, range, hyperlinks}) {
    this.type = type;
    this.shape = shape;
    this.rotation = rotation;
    this.fill = fill;
    this.stroke = stroke;

    if (type === 'shape') {
      if (typeof range === 'string') {
        const decoded = colCache.decode(range);
        this.range = {
          tl: new Anchor(this.worksheet, {col: decoded.left, row: decoded.top}, -1),
          br: new Anchor(this.worksheet, {col: decoded.right, row: decoded.bottom}, 0),
          editAs: 'oneCell',
        };
      } else {
        this.range = {
          tl: new Anchor(this.worksheet, range.tl, 0),
          br: range.br && new Anchor(this.worksheet, range.br, 0),
          ext: range.ext,
          editAs: range.editAs,
          hyperlinks: hyperlinks || range.hyperlinks,
        };
      }
    }
  }
}

Shape.LINE = 'line';
Shape.RECTANGLE = 'rect';
Shape.ROUND_RECTANGLE = 'roundRect';
Shape.ELLIPSE = 'ellipse';
Shape.TRIANGLE = 'triangle';
Shape.RIGHT_ARROW = 'rightArrow';
Shape.DOWN_ARROW = 'downArrow';
Shape.LEFT_BRACE = 'leftBrace';
Shape.RIGHT_BRACE = 'rightBrace';

module.exports = Shape;

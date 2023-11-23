const {parseRange} = require('./drawing-range');

class Shape {
  constructor(worksheet, model) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model() {
    return {
      props: this.props,
      range: {
        tl: this.range.tl.model,
        br: this.range.br && this.range.br.model,
        ext: this.range.ext,
        editAs: this.range.editAs,
      },
    };
  }

  set model({props, range}) {
    this.props = props;
    this.range = parseRange(range, undefined, this.worksheet);
  }
}

// DocumentFormat.OpenXml.Drawing.ShapeTypeValues
Shape.Types = {
  LINE: 'line',
  RECTANGLE: 'rect',
  ROUND_RECTANGLE: 'roundRect',
  ELLIPSE: 'ellipse',
  TRIANGLE: 'triangle',
  RIGHT_ARROW: 'rightArrow',
  DOWN_ARROW: 'downArrow',
  LEFT_BRACE: 'leftBrace',
  RIGHT_BRACE: 'rightBrace',
};

module.exports = Shape;

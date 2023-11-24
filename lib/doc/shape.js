const {parseRange} = require('./drawing-range');

class Shape {
  constructor(worksheet, model) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model() {
    return {
      props: {
        type: this.props.type,
        fill: this.props.fill,
        outline: this.props.outline,
        textBody: this.props.textBody,
      },
      range: {
        tl: this.range.tl.model,
        br: this.range.br && this.range.br.model,
        ext: this.range.ext,
        editAs: this.range.editAs,
      },
    };
  }

  set model({props, range}) {
    this.props = {type: props.type};
    if (props.fill) {
      this.props.fill = props.fill;
    }
    if (props.outline) {
      this.props.outline = props.outline;
    }
    if (props.textBody) {
      this.props.textBody = parseAsTextBody(props.textBody);
    }
    this.range = parseRange(range, undefined, this.worksheet);
  }
}

function parseAsTextBody(input) {
  if (typeof input === 'string') {
    return {
      paragraphs: [parseAsParagraph(input)],
    };
  }
  if (Array.isArray(input)) {
    return {
      paragraphs: input.map(parseAsParagraph),
    };
  }
  return {
    paragraphs: input.paragraphs.map(parseAsParagraph),
  };
}

function parseAsParagraph(input) {
  if (typeof input === 'string') {
    return {
      runs: [parseAsRun(input)],
    };
  }
  if (Array.isArray(input)) {
    return {
      runs: input.map(parseAsRun),
    };
  }
  return {
    runs: input.runs.map(parseAsRun),
  };
}

function parseAsRun(input) {
  if (typeof input === 'string') {
    return {
      text: input,
    };
  }
  return input;
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

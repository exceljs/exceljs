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
        rotation: this.props.rotation,
        horizontalFlip: this.props.horizontalFlip,
        verticalFlip: this.props.verticalFlip,
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
      hyperlinks: this.hyperlinks,
    };
  }

  set model({props, range, hyperlinks}) {
    this.props = {type: props.type};
    if (props.rotation) {
      this.props.rotation = props.rotation;
    }
    if (props.horizontalFlip) {
      this.props.horizontalFlip = props.horizontalFlip;
    }
    if (props.verticalFlip) {
      this.props.verticalFlip = props.verticalFlip;
    }
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
    this.hyperlinks = hyperlinks;
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
  const model = {
    paragraphs: input.paragraphs.map(parseAsParagraph),
  };
  if (input.vertAlign) {
    model.vertAlign = input.vertAlign;
  }
  return model;
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
  const model = {
    runs: input.runs.map(parseAsRun),
  };
  if (input.alignment) {
    model.alignment = input.alignment;
  }
  return model;
}

function parseAsRun(input) {
  if (typeof input === 'string') {
    return {
      text: input,
    };
  }
  const model = {
    text: input.text,
  };
  if (input.font) {
    model.font = input.font;
  }
  return model;
}

module.exports = Shape;

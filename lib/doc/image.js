const {parseRange} = require('./drawing-range');

class Image {
  constructor(worksheet, model) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model() {
    switch (this.type) {
      case 'background':
        return {
          type: this.type,
          imageId: this.imageId,
        };
      case 'image':
        return {
          type: this.type,
          imageId: this.imageId,
          hyperlinks: this.range.hyperlinks,
          range: {
            tl: this.range.tl.model,
            br: this.range.br && this.range.br.model,
            ext: this.range.ext,
            editAs: this.range.editAs,
          },
        };
      default:
        throw new Error('Invalid Image Type');
    }
  }

  set model({type, imageId, range, hyperlinks}) {
    this.type = type;
    this.imageId = imageId;

    if (type === 'image') {
      this.range = parseRange(range, hyperlinks, this.worksheet);
    }
  }
}

module.exports = Image;

const colCache = require('../utils/col-cache');
const Anchor = require('./anchor');

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
          range: {
            tl: this.range.tl.model,
            br: this.range.br && this.range.br.model,
            ext: this.range.ext,
          },
        };
      default:
        throw new Error('Invalid Image Type');
    }
  }

  set model({type, imageId, range}) {
    this.type = type;
    this.imageId = imageId;

    if (type === 'image') {
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
        };
      }
    }
  }
}

module.exports = Image;

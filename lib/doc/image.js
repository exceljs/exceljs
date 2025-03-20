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
          sheetImageId: this.sheetImageId,
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
          sheetImageId: this.sheetImageId,
        };
      default:
        throw new Error('Invalid Image Type');
    }
  }

  set model({type, imageId, range, hyperlinks, sheetImageId}) {
    this.type = type;
    this.imageId = imageId;
    this.sheetImageId = sheetImageId;

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
          hyperlinks: hyperlinks || range.hyperlinks,
        };
      }
    }
  }
}

module.exports = Image;

const colCache = require('../utils/col-cache');
const Anchor = require('./anchor');

module.exports = {
  parseRange: (range, hyperlinks, worksheet) => {
    if (typeof range === 'string') {
      const decoded = colCache.decode(range);
      return {
        tl: new Anchor(worksheet, {col: decoded.left, row: decoded.top}, -1),
        br: new Anchor(worksheet, {col: decoded.right, row: decoded.bottom}, 0),
        editAs: 'oneCell',
      };
    }
    return {
      tl: new Anchor(worksheet, range.tl, 0),
      br: range.br && new Anchor(worksheet, range.br, 0),
      ext: range.ext,
      editAs: range.editAs,
      hyperlinks: hyperlinks || range.hyperlinks,
    };
  },
};

const {Writable} = require('stream');
const saxes = require('saxes');

module.exports = class SAXStream extends Writable {
  initialize() {
    this.sax = new saxes.SaxesParser();
  }

  _write(chunk, _enc, cb) {
    this.sax.write(chunk.toString());
    if (typeof cb === 'function') cb();
  }
};

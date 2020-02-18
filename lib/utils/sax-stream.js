let Stream = require('stream');
const {SaxesParser, EVENTS} = require('saxes');

// Backwards compatibility for earlier node versions and browsers
if (!Stream.Readable || typeof Symbol === 'undefined' || !Stream.Readable.prototype[Symbol.asyncIterator]) {
  Stream = require('readable-stream');
}

module.exports = class SAXStream extends Stream.Transform {
  constructor() {
    super({readableObjectMode: true});
    this._error = null;
    this.saxesParser = new SaxesParser();
    this.saxesParser.on('error', error => {
      this._error = error;
    });
    for (const event of EVENTS) {
      if (event !== 'ready' && event !== 'error' && event !== 'end') {
        this.saxesParser.on(event, value => {
          this.push({event, value});
        });
      }
    }
  }

  _transform(chunk, _encoding, callback) {
    // TODO: Ensure we handle back-pressure!

    this.saxesParser.write(chunk.toString());
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all event handlers
    // have been called
    callback(this._error);
  }

  _final(callback) {
    this.saxesParser.close();
    callback();
  }
};

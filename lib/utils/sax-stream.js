let Stream = require('stream');
const {SaxesParser, EVENTS} = require('saxes');

// Backwards compatibility for earlier node versions and browsers
if (!Stream.Readable || typeof Symbol === 'undefined' || !Stream.Readable.prototype[Symbol.asyncIterator]) {
  Stream = require('readable-stream');
}

const nativeDestroy = Stream.Transform.prototype._destroy; // eslint-disable-line no-underscore-dangle

module.exports = class SAXStream extends Stream.Transform {
  constructor() {
    super({readableObjectMode: true});
    this.saxesParser = new SaxesParser();
    this.saxesParser.on('error', error => {
      this.destroy(error);
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
    this.saxesParser.write(chunk.toString());
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all event handlers
    // have been called
    callback();
  }

  _final(callback) {
    this.saxesParser.close();
    callback();
  }

  _destroy(error, callback) {
    this.saxesParser.close();
    nativeDestroy.call(this, error, callback);
  }
};

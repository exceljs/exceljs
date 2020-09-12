/* eslint-disable node/no-unsupported-features/node-builtins */
// Note: this is included only in browser
if (typeof TextDecoder !== 'undefined' || typeof TextEncoder !== 'undefined') {
  const {Buffer} = require('buffer');
  if (typeof TextEncoder !== 'undefined') {
    const textEncoder = new TextEncoder('utf-8');
    const {from} = Buffer;
    Buffer.from = function(str, encoding) {
      if (typeof str === 'string') {
        if (encoding) {
          encoding = encoding.toLowerCase();
        }
        if (!encoding || encoding === 'utf8' || encoding === 'utf-8') {
          return from.call(this, textEncoder.encode(str).buffer);
        }
      }
      return from.apply(this, arguments);
    };
  }
  if (typeof TextDecoder !== 'undefined') {
    const textDecoder =  new TextDecoder('utf-8');
    const {toString} = Buffer.prototype;
    Buffer.prototype.toString = function(encoding, start, end) {
      if (encoding) {
        encoding = encoding.toLowerCase();
      }
      if (!start && end === undefined &&
        (!encoding || encoding === 'utf8' || encoding === 'utf-8')) {
        return textDecoder.decode(this);
      }
      return toString.apply(this, arguments);
    };
  }
}

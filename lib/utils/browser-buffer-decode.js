// eslint-disable-next-line n/no-unsupported-features/node-builtins
const textDecoder = typeof TextDecoder === 'undefined' ? null : new TextDecoder('utf-8');
const {StringDecoder} = require('string_decoder');

function createTextDecoder() {
  // eslint-disable-next-line n/no-unsupported-features/node-builtins
  if (typeof TextDecoder !== 'undefined') {
    // eslint-disable-next-line n/no-unsupported-features/node-builtins
    const decoder = new TextDecoder('utf-8');
    return {
      decode(chunk) {
        return decoder.decode(chunk, {stream: true});
      },
      flush() {
        return decoder.decode();
      },
    };
  }

  const decoder = new StringDecoder('utf8');
  return {
    decode(chunk) {
      return decoder.write(chunk);
    },
    flush() {
      return decoder.end();
    },
  };
}

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  if (textDecoder) {
    return textDecoder.decode(chunk);
  }
  return chunk.toString();
}

exports.bufferToString = bufferToString;
exports.createTextDecoder = createTextDecoder;

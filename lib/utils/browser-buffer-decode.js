// eslint-disable-next-line node/no-unsupported-features/node-builtins
const textDecoder = typeof TextDecoder === 'undefined' ? null : new TextDecoder('utf-8');
const { StringDecoder } = require('string_decoder');

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  if (StringDecoder) {
    const decoder = new StringDecoder('utf8');
    return decoder.write(chunk)
  }
  if (textDecoder) {
    return textDecoder.decode(chunk);
  }
  return chunk.toString();
}

exports.bufferToString = bufferToString;

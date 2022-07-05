// eslint-disable-next-line node/no-unsupported-features/node-builtins
const textDecoder = typeof TextDecoder === 'undefined' ? null : new TextDecoder('utf-8');
const { StringDecoder } = require('string_decoder');
const decoder =typeof StringDecoder ===  'undefined'? null: new StringDecoder('utf8');

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  if (decoder) {
     return decoder.write(chunk)
  }
  if (textDecoder) {
    return textDecoder.decode(chunk);
  }
  return chunk.toString();
}

exports.bufferToString = bufferToString;

// eslint-disable-next-line node/no-unsupported-features/node-builtins
const textDecoder = typeof TextDecoder === 'undefined' ? null : new TextDecoder('utf-8');

function utf8split(chunk) {
  for (let p = 1; p < chunk.length; p++) {
    const v = chunk[chunk.length - p];
    let len = -1;
    if (v <= 127 && p === 1) {
      // 0b0111_1111
      return [chunk, null];
    }
    if (v <= 191) {
      // 0b1011_1111
      continue;
    }
    if (v <= 223 && p <= 2) {
      // 0b1101_1111
      len = p === 2 ? 0 : p;
    } else if (v <= 239 && p <= 3) {
      // 0b1110_1111
      len = p === 3 ? 0 : p;
    } else if (v <= 247 && p <= 4) {
      // 0b1111_0111
      len = p === 4 ? 0 : p;
    }

    if (len === 0) {
      return [chunk, null];
    }
    if (len > 0) {
      return [chunk.subarray(0, chunk.length - len), chunk.subarray(chunk.length - len)];
    }
    throw new Error('Illegal UTF-8 chunk');
  }
  return [null, chunk];
}

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  const [txtChunk, frag] = utf8split(chunk);
  let txt = '';
  if (txtChunk) {
    if (textDecoder) {
      txt = textDecoder.decode(txtChunk);
    } else {
      txt = txtChunk.toString();
    }
  }
  return {text: txt, chunk: frag};
}

exports.bufferToString = bufferToString;

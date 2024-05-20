const parseSax = verquire('utils/parse-sax');

const utf8Encoder = {
  encode(text) {
    const bytes = [];
    const frags = encodeURIComponent(text).split('%');
    if (frags[0]) {
      bytes.push(...asciiEncode(frags[0]));
    }
    frags.shift();
    frags.forEach(frag => {
      bytes.push(parseInt(frag.substring(0, 2), 16));
      if (frag.length > 2) {
        bytes.push(...asciiEncode(frag));
      }
    });
    return new Uint8Array(bytes);
  },
};

const asciiEncode = ascii => Array.from(ascii).map(c => c.codePointAt(0));

const parse = async chunks => {
  chunks = [[60, 116, 62]].concat(chunks); // <t>
  chunks = chunks.concat([[60, 47, 116, 62]]); // </t>
  chunks = chunks.map(chunk => Buffer.from(chunk));

  const a = [];
  for await (const events of parseSax(chunks)) {
    events.forEach(({eventType, value}) => {
      if (eventType === 'text') {
        a.push(value);
      }
    });
  }
  return a.join('');
};

const utf8FragTest = async text => {
  /* eslint-disable node/no-unsupported-features/node-builtins */
  const encoder =
    typeof TextEncoder === 'undefined' ? utf8Encoder : new TextEncoder('utf-8');
  /* eslint-enable node/no-unsupported-features/node-builtins */

  const bytes = encoder.encode(text);

  // expect(utf8Encoder.encode(text)).to.deep.equal(encoder.encode(text));

  const promises = [];
  promises.push(parse([bytes]));

  // split text to 2 or 3 chunks
  for (let i = 1; i < bytes.length; i++) {
    for (let j = i; j < bytes.length; j++) {
      const chunks = [bytes.subarray(0, i)];
      if (j !== i) {
        chunks.push(bytes.subarray(i, j));
      }
      chunks.push(bytes.subarray(j, bytes.length));
      promises.push(parse(chunks));
    }
  }

  // split same length
  for (let len = 1; len < bytes.length; len++) {
    const chunks = [];
    for (let p = 0; p < bytes.length; p += len) {
      chunks.push(
        bytes.subarray(p, p + len < bytes.length ? p + len : bytes.length)
      );
    }
    promises.push(parse(chunks));
  }
  const results = await Promise.all(promises);
  results.forEach(result => expect(result).to.equal(text));
};

describe('parse-sax', () => {
  it('convert single byte UTF-8 chars properly', async () => {
    await utf8FragTest('abcdefghijklmn');
  });
  it('convert 2 byte UTF-8 chars properly', async () => {
    await utf8FragTest('αβγδεζηθικ');
  });
  it('convert 3 byte UTF-8 chars properly', async () => {
    await utf8FragTest('あいうえお');
  });
  it('convert 4 byte UTF-8 chars properly', async () => {
    await utf8FragTest('😀😁😂😃😄');
  });
});

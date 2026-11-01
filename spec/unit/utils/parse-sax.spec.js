const parseSax = verquire('utils/parse-sax');

describe('parse-sax', () => {
  it('decodes UTF-8 characters split across input chunks', async () => {
    const firstCharacter = Buffer.from('中');
    const chunks = [
      Buffer.concat([Buffer.from('<root>'), firstCharacter.slice(0, 1)]),
      Buffer.concat([firstCharacter.slice(1), Buffer.from('文</root>')]),
    ];
    const text = [];

    async function* input() {
      yield* chunks;
    }

    for await (const events of parseSax(input())) {
      events.forEach(event => {
        if (event.eventType === 'text') {
          text.push(event.value);
        }
      });
    }

    expect(text.join('')).to.equal('中文');
  });
});

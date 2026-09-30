module.exports = async function* iterateStream(stream) {
  const contents = [];
  stream.on('data', data => contents.push(data));

  let resolveStreamEndedPromise;
  const streamEndedPromise = new Promise(resolve => (resolveStreamEndedPromise = resolve));

  let ended = false;
  stream.on('end', () => {
    ended = true;
    resolveStreamEndedPromise();
  });

  let error = false;
  stream.on('error', err => {
    error = err;
    resolveStreamEndedPromise();
  });

  // Some streams emit 'end' themselves while chunks are still buffered on the readable
  // side (unzipper does it on writable 'finish'), so keep reading until the buffer is empty.
  while (!ended || contents.length > 0 || stream.readableLength > 0) {
    if (contents.length === 0) {
      stream.resume();
      // eslint-disable-next-line no-await-in-loop
      await (ended ? once(stream, 'data') : Promise.race([once(stream, 'data'), streamEndedPromise]));
    } else {
      stream.pause();
      const data = contents.shift();
      yield data;
    }
    if (error) throw error;
  }
  resolveStreamEndedPromise();
};

function once(eventEmitter, type) {
  // TODO: Use require('events').once when node v10 is dropped
  return new Promise(resolve => {
    let fired = false;
    const handler = () => {
      if (!fired) {
        fired = true;
        eventEmitter.removeListener(type, handler);
        resolve();
      }
    };
    eventEmitter.addListener(type, handler);
  });
}

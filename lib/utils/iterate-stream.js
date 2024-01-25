module.exports = async function* iterateStream(stream) {
  const contents = [];
  stream.on('data', data => contents.push(data));

  let error = false;
  stream.on('error', err => {
    error = err;
  });

  while (!stream.closed || contents.length) {
    if (contents.length === 0) {
      const streamEndedPromise = new Promise(resolve => {
        stream.resume().on('end', () => {
          resolve();
        });
      });
      // eslint-disable-next-line no-await-in-loop
      await Promise.race([once(stream, 'data'), streamEndedPromise]);
    } else {
      stream.pause();
      const data = contents.shift();
      yield data;
    }
    if (error) throw error;
  }
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

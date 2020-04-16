module.exports = async function* iterateStream(stream) {
  const contents = [];
  stream.on('data', data => contents.push(data));

  let finished = false;
  stream.on('finish', () => (finished = true));
  const streamFinishedPromise = once(stream, 'finish');

  let error = false;
  stream.on('error', err => (error = err));
  const errorPromise = once(stream, 'error');

  while (!finished || contents.length > 0) {
    if (error) throw error;
    if (contents.length === 0) {
      // eslint-disable-next-line no-await-in-loop
      await Promise.race([once(stream, 'data'), errorPromise, streamFinishedPromise]);
    } else {
      const data = contents.shift();
      yield data;
    }
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

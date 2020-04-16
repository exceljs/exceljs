module.exports = async function* iterateStream(entry) {
  const contents = [];
  entry.on('data', data => contents.push(data));

  let finished = false;
  entry.on('finish', () => {
    finished = true;
  });
  const entryFinishedPromise = once(entry, 'finish');

  while (!finished || contents.length > 0) {
    if (contents.length === 0) {
      // eslint-disable-next-line no-await-in-loop
      await Promise.race([once(entry, 'data'), entryFinishedPromise]);
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

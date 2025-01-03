const { once } = require('events');

module.exports = async function* iterateStream(stream) {
  const contents = [];
  let ended = false;
  let error = null;

  const onData = data => contents.push(data);
  const onEnd = () => { ended = true; };
  const onError = err => { error = err; };

  stream.on('data', onData);
  stream.on('end', onEnd);
  stream.on('error', onError);

  try {
    while (!ended || contents.length > 0) {
      // first proccess thje data
      if (contents.length > 0) {
        const chunk = contents.shift();
        yield chunk;
        continue;
      }

      // wait for new data or end 
      if (!ended) {
        await Promise.race([once(stream, 'data'), once(stream, 'end')]);
      }

      if (error) throw error;
    }
  } finally {
    // cleanup to prevent leaks
    stream.off('data', onData);
    stream.off('end', onEnd);
    stream.off('error', onError);
  }
};

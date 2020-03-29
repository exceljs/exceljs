const {SaxesParser} = require('saxes');

module.exports = async function*(stream, eventTypes) {
  const saxesParser = new SaxesParser();
  let error;
  saxesParser.on('error', err => {
    error = err;
  });
  let events = [];
  for (const eventType of ['opentag', 'text', 'closetag']) {
    saxesParser.on(eventType, value => {
      events.push({eventType, value});
    });
  }
  for await (const chunk of stream) {
    saxesParser.write(chunk.toString());
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all events have been emitted
    if (error) throw error;
    // As a performance optimization, we gather all events instead of passing
    // them one by one, which would cause each event to go through the event queue
    yield events;
    events = [];
  }
};

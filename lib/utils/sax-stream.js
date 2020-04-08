const Stream = require('readable-stream');
const {SaxesParser} = require('saxes');

module.exports = class SAXStream extends Stream.Transform {
  constructor(eventTypes) {
    super({readableObjectMode: true});
    this.events = [];
    this.saxesParser = new SaxesParser();
    this.saxesParser.on('error', error => {
      this.destroy(error);
    });
    for (const eventType of eventTypes) {
      if (eventType !== 'ready' && eventType !== 'error' && eventType !== 'end') {
        this.saxesParser.on(eventType, value => {
          this.events.push({eventType, value});
        });
      }
    }
  }

  _transform(chunk, _encoding, callback) {
    this.saxesParser.write(chunk.toString());
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all events have been emitted
    const {events} = this;
    this.events = [];
    // As a performance optimization, we gather all events instead of passing
    // them one by one, which would cause each event to go through the event queue
    callback(null, events);
  }

  _final(callback) {
    this.saxesParser.close();
    callback();
  }
};

const events = require('events');

// =============================================================================
// AutoDrain - kind of /dev/null
class AutoDrain extends events.EventEmitter {
  write(chunk) {
    this.emit('data', chunk);
  }

  end() {
    this.emit('end');
  }
}

module.exports = AutoDrain;

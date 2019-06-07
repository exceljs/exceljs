const events = require('events');
const PromiseLib = require('./promise');

const utils = require('./utils');

// =============================================================================
// FlowControl - Used to slow down streaming to manageable speed
// Implements a subset of Stream.Duplex: pipe() and write()
class FlowControl extends events.EventEmitter {
  constructor(options) {
    super();

    this.options = options = options || {};

    // Buffer queue
    this.queue = [];

    // Consumer streams
    this.pipes = [];

    // Down-stream flow-control instances
    this.children = [];

    // Up-stream flow-control instances
    this.parent = options.parent;

    // Ensure we don't flush more than once at a time
    this.flushing = false;

    // determine timeout for flow control delays
    if (options.gc) {
      const {gc} = options;
      if (gc.getTimeout) {
        this.getTimeout = gc.getTimeout;
      } else {
        // heap size below which we don't bother delaying
        const threshold = gc.threshold !== undefined ? gc.threshold : 150000000;
        // convert from heapsize to ms timeout
        const divisor = gc.divisor !== undefined ? gc.divisor : 500000;
        this.getTimeout = function() {
          const memory = process.memoryUsage();
          const heapSize = memory.heapTotal;
          return heapSize < threshold ? 0 : Math.floor(heapSize / divisor);
        };
      }
    } else {
      this.getTimeout = null;
    }
  }

  get name() {
    return ['FlowControl', this.parent ? 'Child' : 'Root', this.corked ? 'corked' : 'open'].join(' ');
  }

  get corked() {
    // We remain corked while we have children and at least one has data to consume
    return this.children.length > 0 && this.children.some(child => child.queue && child.queue.length);
  }

  get stem() {
    // the decision to stem the incoming data depends on whether the children are corked
    // and how many buffers we have backed up
    return this.corked || !this.queue || this.queue.length > 2;
  }

  _write(dst, data, encoding) {
    // Write to a single destination and return a promise

    return new PromiseLib.Promise((resolve, reject) => {
      dst.write(data, encoding, error => {
        if (error) {
          reject(error);
        } else {
          resolve();
        }
      });
    });
  }

  _pipe(chunk) {
    // Write chunk to all pipes. A chunk with no data is the end
    const promises = [];
    this.pipes.forEach(pipe => {
      if (chunk.data) {
        if (pipe.sync) {
          pipe.stream.write(chunk.data, chunk.encoding);
        } else {
          promises.push(this._write(pipe.stream, chunk.data, chunk.encoding));
        }
      } else {
        pipe.stream.end();
      }
    });
    if (!promises.length) {
      promises.push(PromiseLib.Promise.resolve());
    }
    return PromiseLib.Promise.all(promises).then(() => {
      try {
        chunk.callback();
      } catch (e) {
        // quietly ignore
      }
    });
  }

  _animate() {
    let count = 0;
    const seq = ['|', '/', '-', '\\'];
    const cr = '\u001b[0G'; // was '\033[0G'
    return setInterval(() => {
      process.stdout.write(seq[count++ % 4] + cr);
    }, 100);
  }

  _delay() {
    // in certain situations it may be useful to delay processing (e.g. for GC)
    const timeout = this.getTimeout && this.getTimeout();
    if (timeout) {
      return new PromiseLib.Promise(resolve => {
        const anime = this._animate();
        setTimeout(() => {
          clearInterval(anime);
          resolve();
        }, timeout);
      });
    }
    return PromiseLib.Promise.resolve();
  }

  _flush() {
    // If/while not corked and we have buffers to send, send them
    if (this.queue && !this.flushing && !this.corked) {
      if (this.queue.length) {
        this.flushing = true;
        this._delay()
          .then(() => this._pipe(this.queue.shift()))
          .then(() => {
            setImmediate(() => {
              this.flushing = false;
              this._flush();
            });
          });
      }

      if (!this.stem) {
        // Signal up-stream that we're ready for more data
        this.emit('drain');
      }
    }
  }

  write(data, encoding, callback) {
    // Called by up-stream pipe
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = 'utf8';
    }
    callback = callback || utils.nop;

    if (!this.queue) {
      throw new Error('Cannot write to stream after end');
    }

    // Always queue chunks and then flush
    this.queue.push({
      data,
      encoding,
      callback,
    });
    this._flush();

    // restrict further incoming data if we have backed up buffers or
    // the children are still busy
    const stemFlow = this.corked || this.queue.length > 3;
    return !stemFlow;
  }

  end() {
    // Signal from up-stream
    this.queue.push({
      callback: () => {
        this.queue = null;
        this.emit('finish');
      },
    });
    this._flush();
  }

  pipe(stream, options) {
    options = options || {};

    // some streams don't call callbacks
    const sync = options.sync || false;

    this.pipes.push({
      stream,
      sync,
    });
  }

  unpipe(stream) {
    this.pipes = this.pipes.filter(pipe => pipe.stream !== stream);
  }

  createChild() {
    // Create a down-stream flow-control
    const options = Object.assign({parent: this}, this.options);
    const child = new FlowControl(options);
    this.children.push(child);

    child.on('drain', () => {
      // a child is ready for more
      this._flush();
    });

    child.on('finish', () => {
      // One child has finished its stream. Remove it and continue
      this.children = this.children.filter(item => item !== child);
      this._flush();
    });

    return child;
  }
}

module.exports = FlowControl;

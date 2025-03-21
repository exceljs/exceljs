const { EventEmitter } = require('events');

class LineBuffer extends EventEmitter {
  constructor(options = {}) {
    super();

    // Encoding, default to UTF-8
    this.encoding = options.encoding || 'utf8';
    
    // Initialize buffer and line queue
    this.buffer = '';
    this.queue = [];
    this.lineCount = 0;

    // Cork/uncork state and buffer size limit
    this.corked = false;
    this.maxBufferSize = options.maxBufferSize || 1024 * 1024;  // Default to 1 MB
    this.maxLineCount = options.maxLineCount || Infinity; // Default to no line limit
    
    // Transform function for line processing
    this.transformFn = options.transform || null;

    // Pause and resume functionality
    this._paused = false;

    // Custom line delimiter (default is \n)
    this.lineDelimiter = options.lineDelimiter || /\r?\n/;

    // Logging flag
    this.logging = options.logging || false;

    // Progress tracking
    this.progressCallback = options.progressCallback || null;

    // Optional: Flush after a certain number of lines
    this.flushLineCount = options.flushLineCount || Infinity;
  }

  // Events:
  // 'line': Emits a processed line
  // 'done': Emits when all lines are emitted
  // 'error': Emits if any error occurs during processing
  // 'drain': Emits when the buffer has been uncorked and is ready for more data
  
  write(chunk) {
    if (this._paused) {
      return false;  // If paused, don't process the chunk
    }

    // Ensure the chunk is treated as a string (use the encoding)
    chunk = chunk.toString(this.encoding);

    // Add the chunk to the buffer
    this.buffer += chunk;

    // Process the buffer and split it by custom line delimiter
    let lines = this.buffer.split(this.lineDelimiter);
    
    // The last part might not be a full line, so we'll keep it as the buffer
    this.buffer = lines.pop();

    // Emit each line or queue it if corked
    lines.forEach(line => {
      if (this.corked) {
        this.queue.push(line);
      } else {
        if (this.transformFn) {
          line = this.transformFn(line);
        }
        this.lineCount++;
        this.emit('line', line);

        // Track progress if callback is provided
        if (this.progressCallback) {
          this.progressCallback(this.lineCount);
        }

        // Auto-flush after certain number of lines
        if (this.lineCount >= this.flushLineCount) {
          this._flush();
        }
      }
    });

    // If buffer exceeds size limit, automatically flush
    if (this.buffer.length > this.maxBufferSize) {
      this._flush();
    }

    return !this.corked;
  }

  cork() {
    this.corked = true;
  }

  uncork() {
    this.corked = false;
    this._flush();
    this.emit('drain');
  }

  setDefaultEncoding(encoding) {
    this.encoding = encoding;
  }

  end(callback) {
    if (this.buffer) {
      this.emit('line', this.buffer);
      this.buffer = '';
    }
    this.emit('done');
    
    if (callback) {
      callback();  // Allow a custom callback when processing ends
    }
  }

  _flush() {
    if (!this.corked && this.queue.length > 0) {
      this.queue.forEach(line => {
        if (this.transformFn) {
          line = this.transformFn(line);
        }
        this.emit('line', line);
      });
      this.queue = [];
    }
  }

  // Pause the stream (won't process any more data)
  pause() {
    this._paused = true;
  }

  // Resume processing the stream
  resume() {
    this._paused = false;
    if (this.buffer) {
      this.write(this.buffer);
    }
    this._flush();
  }

  // Emit an error if needed
  emitError(errorMessage) {
    this.emit('error', new Error(errorMessage));
  }

  // Log activity (optional)
  log(message) {
    if (this.logging) {
      console.log(`[LineBuffer] ${message}`);
    }
  }

  // Get current buffer statistics
  getStats() {
    return {
      currentBufferSize: this.buffer.length,
      lineCount: this.lineCount,
      queuedLines: this.queue.length,
      corked: this.corked
    };
  }
}

module.exports = LineBuffer;

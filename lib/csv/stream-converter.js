// =======================================================================================================
// StreamConverter
//
// convert between encoding schemes in a stream
// Improved version with added features and error handling

let jconv;

class StreamConverter {
  constructor(inner, options = {}) {
    this.inner = inner;

    // Set encoding options and validate them
    this.innerEncoding = (options.innerEncoding || 'UTF8').toUpperCase();
    this.outerEncoding = (options.outerEncoding || 'UTF8').toUpperCase();

    this.innerBOM = options.innerBOM || null;
    this.outerBOM = options.outerBOM || null;

    this.writeStarted = false;

    // Option to handle encoding detection
    this.detectEncoding = options.detectEncoding || false;
    this.customErrorHandler = options.customErrorHandler || null;
  }

  // Error handling for invalid encoding type
  validateEncoding(encoding) {
    const validEncodings = ['UTF8', 'UTF16', 'ASCII', 'BASE64', 'UTF32', 'ISO-8859-1', 'WINDOWS-1252']; // Add more as needed
    if (!validEncodings.includes(encoding.toUpperCase())) {
      throw new Error(`Invalid encoding: ${encoding}`);
    }
  }

  // Converts incoming data from outer encoding to inner encoding
  convertInwards(data) {
    if (data) {
      if (typeof data === 'string') {
        data = Buffer.from(data, this.outerEncoding);
      }

      if (this.innerEncoding !== this.outerEncoding) {
        data = jconv.convert(data, this.outerEncoding, this.innerEncoding);
      }
    }

    return data;
  }

  // Converts outgoing data from inner encoding to outer encoding
  convertOutwards(data) {
    if (typeof data === 'string') {
      data = Buffer.from(data, this.innerEncoding);
    }

    if (this.innerEncoding !== this.outerEncoding) {
      data = jconv.convert(data, this.innerEncoding, this.outerEncoding);
    }

    return data;
  }

  // Read method: Handles incoming data if implemented
  read(size) {
    return new Promise((resolve, reject) => {
      try {
        this.inner.read(size, (err, data) => {
          if (err) reject(err);
          resolve(this.convertOutwards(data));
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  // Add a listener for a specific event
  addListener(event, handler) {
    this.inner.addListener(event, handler);
  }

  // Remove a listener for a specific event
  removeListener(event, handler) {
    this.inner.removeListener(event, handler);
  }

  // Write method: Handles encoding conversion and writes data
  write(data, encoding, callback) {
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = undefined;
    }

    if (!this.writeStarted) {
      // Write BOM if inner encoding has BOM
      if (this.innerBOM) {
        this.inner.write(this.innerBOM);
      }

      // Remove BOM if outer encoding has BOM
      if (this.outerBOM) {
        if (data.length <= this.outerBOM.length) {
          if (callback) {
            callback();
          }
          return;
        }
        const bomless = Buffer.alloc(data.length - this.outerBOM.length);
        data.copy(bomless, 0, this.outerBOM.length, data.length);
        data = bomless;
      }

      this.writeStarted = true;
    }

    // Handle encoding detection and conversion
    if (this.detectEncoding) {
      this.validateEncoding(this.outerEncoding);
      this.validateEncoding(this.innerEncoding);
    }

    try {
      this.inner.write(
        this.convertInwards(data),
        encoding ? this.innerEncoding : undefined,
        callback
      );
    } catch (err) {
      if (this.customErrorHandler) {
        this.customErrorHandler(err);
      } else {
        throw err;
      }
    }
  }

  // Pipe method: Supports chaining streams with conversion
  pipe(destination, options) {
    const reverseConverter = new StreamConverter(destination, {
      innerEncoding: this.outerEncoding,
      outerEncoding: this.innerEncoding,
      innerBOM: this.outerBOM,
      outerBOM: this.innerBOM,
      detectEncoding: this.detectEncoding,
      customErrorHandler: this.customErrorHandler,
    });

    this.inner.pipe(reverseConverter, options);
  }

  // Close the stream
  close() {
    this.inner.close();
  }

  // Handles incoming data and converts before passing to the callback
  on(type, callback) {
    switch (type) {
      case 'data':
        this.inner.on('data', chunk => {
          callback(this.convertOutwards(chunk));
        });
        return this;
      case 'end':
        this.inner.on('end', () => {
          callback();
        });
        return this;
      case 'error':
        this.inner.on('error', err => {
          if (this.customErrorHandler) {
            this.customErrorHandler(err);
          } else {
            throw err;
          }
        });
        return this;
      default:
        this.inner.on(type, callback);
        return this;
    }
  }

  once(type, callback) {
    this.inner.once(type, callback);
  }

  // End the stream with data and optional encoding
  end(chunk, encoding, callback) {
    try {
      this.inner.end(this.convertInwards(chunk), this.innerEncoding, callback);
    } catch (err) {
      if (this.customErrorHandler) {
        this.customErrorHandler(err);
      } else {
        throw err;
      }
    }
  }

  // Emit events for the inner stream
  emit(type, value) {
    this.inner.emit(type, value);
  }

  // Added method to detect BOM in incoming data
  detectBOM(data) {
    if (data.slice(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf]))) {
      return 'UTF8';
    }
    return null;
  }

  // Optional transformation functionality
  transformData(chunk, transformer) {
    return transformer ? transformer(chunk) : chunk;
  }

  // Buffering to handle large data chunks more efficiently
  bufferData(data, maxSize = 1024 * 1024) {
    let buffer = [];
    let bufferSize = 0;

    for (let i = 0; i < data.length; i++) {
      buffer.push(data[i]);
      bufferSize++;

      // If buffer exceeds max size, flush it
      if (bufferSize >= maxSize) {
        this.write(Buffer.concat(buffer), null, () => {
          buffer = [];
          bufferSize = 0;
        });
      }
    }

    // Write any remaining data in the buffer
    if (buffer.length > 0) {
      this.write(Buffer.concat(buffer));
    }
  }
}

module.exports = StreamConverter;

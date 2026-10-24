// This fork runs on Node 18+ (global fetch, fs.promises). Media is added to
// the zip one entry at a time, in order, so awaiting inside the loop is intended.
/* eslint-disable no-await-in-loop, node/no-unsupported-features/node-builtins */
/* global fetch, AbortSignal */
const fs = require('fs');
const path = require('path');

const DEFAULT_PLACEHOLDER = path.join(__dirname, 'media', 'placeholder.png');

const DEFAULTS = {
  placeholder: DEFAULT_PLACEHOLDER,
  onError: null,
  concurrency: 8,
  timeout: 30000,
  maxBytes: 20 * 1024 * 1024,
  headers: {},
};

// Leading bytes of the image formats Excel can display. Media with other
// extensions (emf, wmf, ...) is passed through without checking.
const SIGNATURES = {
  png: [[0x89, 0x50, 0x4e, 0x47]],
  jpeg: [[0xff, 0xd8, 0xff]],
  jpg: [[0xff, 0xd8, 0xff]],
  gif: [[0x47, 0x49, 0x46, 0x38]],
  bmp: [[0x42, 0x4d]],
  tif: [
    [0x49, 0x49, 0x2a, 0x00],
    [0x4d, 0x4d, 0x00, 0x2a],
  ],
  tiff: [
    [0x49, 0x49, 0x2a, 0x00],
    [0x4d, 0x4d, 0x00, 0x2a],
  ],
};
const ANY_IMAGE = Object.values(SIGNATURES).reduce((all, sigs) => all.concat(sigs), []);

function isValidImage(header, extension) {
  if (!SIGNATURES[extension]) {
    return true;
  }
  // The extension says what the caller expects, but any known image format
  // displays fine (callers often label PNGs as jpg), so accept all of them.
  return ANY_IMAGE.some(sig => header.length >= sig.length && sig.every((byte, i) => header[i] === byte));
}

function invalidImageError(source) {
  const error = new Error(`Not a valid image: ${source}`);
  error.code = 'EINVALIDIMAGE';
  return error;
}

async function readHeader(filename) {
  const handle = await fs.promises.open(filename, 'r');
  try {
    const {buffer, bytesRead} = await handle.read(Buffer.alloc(8), 0, 8, 0);
    return buffer.subarray(0, bytesRead);
  } finally {
    await handle.close();
  }
}

async function downloadImage(url, options) {
  const res = await fetch(url, {
    headers: options.headers,
    signal: AbortSignal.timeout(options.timeout),
    redirect: 'follow',
  });
  if (!res.ok) {
    const error = new Error(`HTTP ${res.status} for ${url}`);
    error.statusCode = res.status;
    throw error;
  }
  const tooLarge = () => new Error(`Image larger than ${options.maxBytes} bytes: ${url}`);
  if (Number(res.headers.get('content-length')) > options.maxBytes) {
    throw tooLarge();
  }
  const buffer = Buffer.from(await res.arrayBuffer());
  if (buffer.length > options.maxBytes) {
    throw tooLarge();
  }
  return buffer;
}

async function fetchImage(url, options) {
  if (typeof fetch !== 'function') {
    throw new Error('Images added by url need Node 18 or later (global fetch)');
  }
  try {
    return await downloadImage(url, options);
  } catch (error) {
    // the timeout covers the whole download, headers and body
    if (error.name === 'TimeoutError') {
      const timeout = new Error(`Timed out after ${options.timeout} ms: ${url}`);
      timeout.code = 'ETIMEDOUT';
      throw timeout;
    }
    throw error;
  }
}

// Adds workbook media to an archiver zip.
//
// - filename images are checked, then streamed from disk lazily (archiver.file)
// - url images are downloaded `concurrency` at a time ahead of the writer, and
//   at most `concurrency` downloaded images wait in the zip queue at once
// - a missing, unreadable, failed or invalid image is replaced by the
//   placeholder (pass `placeholder: false` to fail the write instead)
class MediaSource {
  constructor(options) {
    this.options = Object.assign({}, DEFAULTS, options);
    this.placeholder = null;
  }

  async _getPlaceholder() {
    if (!this.placeholder) {
      const {placeholder} = this.options;
      this.placeholder = Buffer.isBuffer(placeholder) ? placeholder : await fs.promises.readFile(placeholder);
    }
    return this.placeholder;
  }

  // Returns the placeholder Buffer, or rethrows when placeholders are off.
  async _fallback(error, medium) {
    if (this.options.placeholder === false) {
      throw error;
    }
    if (this.options.onError) {
      this.options.onError(error, medium);
    }
    return this._getPlaceholder();
  }

  async _resolveFile(medium) {
    try {
      if (!isValidImage(await readHeader(medium.filename), medium.extension)) {
        throw invalidImageError(medium.filename);
      }
      return {filename: medium.filename};
    } catch (error) {
      return {buffer: await this._fallback(error, medium)};
    }
  }

  async _resolveBuffer(buffer, medium) {
    if (isValidImage(buffer, medium.extension)) {
      return {buffer};
    }
    return {buffer: await this._fallback(invalidImageError(medium.url || 'buffer'), medium)};
  }

  async _resolveUrl(download, medium) {
    const {buffer, error} = await download;
    if (error) {
      return {buffer: await this._fallback(error, medium)};
    }
    return this._resolveBuffer(buffer, medium);
  }

  async addAll(zip, media) {
    const {concurrency} = this.options;
    const downloads = new Map();
    let nextDownload = 0;
    const startDownloads = upTo => {
      for (; nextDownload < media.length && nextDownload <= upTo; nextDownload++) {
        const medium = media[nextDownload];
        if (medium.type === 'image' && medium.url) {
          // never rejects, so a failure waiting its turn is not "unhandled"
          downloads.set(
            nextDownload,
            fetchImage(medium.url, this.options).then(
              buffer => ({buffer}),
              error => ({error})
            )
          );
        }
      }
    };

    // names of downloaded images appended but not yet written by archiver
    const queued = new Set();
    let waiter = null;
    let zipError = null;
    const entryListener = entry => {
      if (queued.delete(entry.name) && waiter) {
        waiter.resolve();
      }
    };
    const errorListener = error => {
      zipError = error;
      if (waiter) {
        waiter.reject(error);
      }
    };
    zip.on('entry', entryListener);
    zip.on('error', errorListener);
    const waitForWrite = () =>
      new Promise((resolve, reject) => {
        waiter = {resolve, reject};
      });

    try {
      for (let i = 0; i < media.length; i++) {
        startDownloads(i + concurrency);
        const medium = media[i];
        if (medium.type !== 'image') {
          throw new Error('Unsupported media');
        }
        const name = `xl/media/${medium.name}.${medium.extension}`;

        let source;
        if (medium.url) {
          source = await this._resolveUrl(downloads.get(i), medium);
          downloads.delete(i);
        } else if (medium.filename) {
          source = await this._resolveFile(medium);
        } else if (medium.buffer) {
          source = await this._resolveBuffer(medium.buffer, medium);
        } else if (medium.base64) {
          const dataimg64 = medium.base64;
          const content = dataimg64.substring(dataimg64.indexOf(',') + 1);
          source = await this._resolveBuffer(Buffer.from(content, 'base64'), medium);
        } else {
          throw new Error('Unsupported media');
        }

        if (source.filename) {
          zip.file(source.filename, {name});
          continue;
        }
        zip.append(source.buffer, {name});
        if (medium.url) {
          queued.add(name);
          while (queued.size > concurrency) {
            if (zipError) {
              throw zipError;
            }
            await waitForWrite();
            waiter = null;
          }
        }
      }
    } finally {
      zip.removeListener('entry', entryListener);
      zip.removeListener('error', errorListener);
    }
  }
}

MediaSource.DEFAULT_PLACEHOLDER = DEFAULT_PLACEHOLDER;

module.exports = MediaSource;

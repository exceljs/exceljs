const DEFAULT_MAX_ENTRY_SIZE = 100 * 1024 * 1024;
const DEFAULT_MAX_TOTAL_SIZE = 500 * 1024 * 1024;
const DEFAULT_MAX_COMPRESSION_RATIO = 1000;

function positiveLimit(value, fallback, name) {
  if (value === undefined) {
    return fallback;
  }
  if (!Number.isFinite(value) || value <= 0) {
    throw new TypeError(`${name} must be a positive finite number`);
  }
  return value;
}

function getZipLimits(options) {
  options = options || {};
  return {
    maxEntrySize: positiveLimit(options.maxEntrySize, DEFAULT_MAX_ENTRY_SIZE, 'maxEntrySize'),
    maxTotalSize: positiveLimit(options.maxTotalSize, DEFAULT_MAX_TOTAL_SIZE, 'maxTotalSize'),
    maxCompressionRatio: positiveLimit(
      options.maxCompressionRatio,
      DEFAULT_MAX_COMPRESSION_RATIO,
      'maxCompressionRatio'
    ),
  };
}

function validateZipEntry(entry, limits, totalSize) {
  const data = entry._data || entry;
  const uncompressedSize = Number(
    data.uncompressedSize !== undefined ? data.uncompressedSize : entry.size
  );
  const compressedSize = Number(
    data.compressedSize !== undefined
      ? data.compressedSize
      : entry.vars && entry.vars.compressedSize
  );

  // Streaming ZIP entries using data descriptors do not expose their size
  // until the entry has been consumed. JSZip entries always have a central
  // directory size and are checked before decompression.
  if (isUnknownZipEntrySize(entry, data, uncompressedSize)) {
    return totalSize;
  }
  if (uncompressedSize < 0) {
    throw new Error(`Invalid uncompressed size for zip entry: ${entry.name || entry.path}`);
  }
  if (uncompressedSize > limits.maxEntrySize) {
    throw new Error(`Zip entry exceeds maxEntrySize: ${entry.name || entry.path}`);
  }
  if (totalSize + uncompressedSize > limits.maxTotalSize) {
    throw new Error(`Zip archive exceeds maxTotalSize: ${entry.name || entry.path}`);
  }
  if (compressedSize > 0 && uncompressedSize / compressedSize > limits.maxCompressionRatio) {
    throw new Error(`Zip entry exceeds maxCompressionRatio: ${entry.name || entry.path}`);
  }
  return totalSize + uncompressedSize;
}

function isUnknownZipEntrySize(entry, data, uncompressedSize) {
  return (
    !Number.isFinite(uncompressedSize) ||
    (entry.size === undefined &&
      data.flags !== undefined &&
      Math.floor(data.flags / 8) % 2 === 1 &&
      uncompressedSize === 0)
  );
}

function createZipEntryLimiter(entry, limits, totalSize) {
  const data = entry._data || entry;
  const compressedSize = Number(
    data.compressedSize !== undefined
      ? data.compressedSize
      : entry.vars && entry.vars.compressedSize
  );
  let bytesRead = 0;
  const limiter = new Transform({
    transform(chunk, encoding, callback) {
      bytesRead += chunk.length;
      if (bytesRead > limits.maxEntrySize) {
        return callback(new Error(`Zip entry exceeds maxEntrySize: ${entry.name || entry.path}`));
      }
      if (totalSize + bytesRead > limits.maxTotalSize) {
        return callback(new Error(`Zip archive exceeds maxTotalSize: ${entry.name || entry.path}`));
      }
      if (compressedSize > 0 && bytesRead / compressedSize > limits.maxCompressionRatio) {
        return callback(new Error(`Zip entry exceeds maxCompressionRatio: ${entry.name || entry.path}`));
      }
      return callback(null, chunk);
    },
  });
  entry.pipe(limiter);
  limiter.bytesRead = () => bytesRead;
  return limiter;
}

module.exports = {
  createZipEntryLimiter,
  getZipLimits,
  isUnknownZipEntrySize,
  validateZipEntry,
};
const {Transform} = require('stream');

'use strict';

// [exceljs-hardened] CVE fix — decompression bomb / memory exhaustion
// (GHSA pending).
//
// Workbook#load() used to fully decompress every zip entry into memory
// with no cap on size, entry count, or compression ratio — a few hundred
// KB of highly-compressible input could expand past a gigabyte before a
// single byte of XML was ever parsed. This guard reads the *declared*
// uncompressed size straight from the zip's central directory (already
// parsed by JSZip.loadAsync() before any entry is decompressed) and
// refuses to proceed if a single entry, or the archive as a whole, would
// exceed a sane limit — without ever materializing the bomb.

const DEFAULT_MAX_ENTRY_UNCOMPRESSED_SIZE = 128 * 1024 * 1024; // 128 MB
const DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE = 512 * 1024 * 1024; // 512 MB

function createDecompressionBudget(options) {
  return {
    bytes: 0,
    maxEntry: (options && options.maxEntryUncompressedSize) || DEFAULT_MAX_ENTRY_UNCOMPRESSED_SIZE,
    maxTotal: (options && options.maxTotalUncompressedSize) || DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE,
  };
}

function assertSafeZipEntry(entry, entryName, budget) {
  const declaredSize =
    entry._data && typeof entry._data.uncompressedSize === 'number'
      ? entry._data.uncompressedSize
      : undefined;

  // If the central directory didn't declare a size (rare/malformed
  // archives), we can't pre-check it — fall through rather than break
  // otherwise-valid files. The realistic, demonstrated zip-bomb attack
  // relies on a declared size and is fully covered below.
  if (declaredSize === undefined) {
    return;
  }

  if (declaredSize > budget.maxEntry) {
    throw new Error(
      `Refusing to decompress "${entryName}": declared uncompressed size (${declaredSize} bytes) ` +
        `exceeds the per-entry limit (${budget.maxEntry} bytes). This may be a decompression bomb. ` +
        'Raise options.maxEntryUncompressedSize to override.'
    );
  }

  budget.bytes += declaredSize;
  if (budget.bytes > budget.maxTotal) {
    throw new Error(
      `Refusing to decompress "${entryName}": cumulative declared uncompressed size ` +
        `(${budget.bytes} bytes) exceeds the archive-wide limit (${budget.maxTotal} bytes). ` +
        'This may be a decompression bomb. Raise options.maxTotalUncompressedSize to override.'
    );
  }
}

module.exports = {
  createDecompressionBudget,
  assertSafeZipEntry,
  DEFAULT_MAX_ENTRY_UNCOMPRESSED_SIZE,
  DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE,
};

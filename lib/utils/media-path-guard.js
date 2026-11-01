const path = require('path');

function assertSafeMediaPath(filename) {
  if (typeof filename !== 'string' || filename.length === 0) {
    throw new TypeError('Image filename must be a non-empty string');
  }
  if (filename.includes('\0')) {
    throw new Error('Image filename contains a null byte');
  }

  // Check both separators so a path validated on one platform cannot become
  // a traversal path when it is later used on another platform.
  const segments = filename.replace(/\\/g, '/').split('/');
  if (!path.isAbsolute(filename) && segments.includes('..')) {
    throw new Error(`Image filename must not contain path traversal: ${filename}`);
  }
  return filename;
}

function safeJoin(baseDir, userPath) {
  assertSafeMediaPath(userPath);
  const base = path.resolve(baseDir);
  const resolved = path.resolve(base, userPath);
  if (resolved !== base && !resolved.startsWith(`${base}${path.sep}`)) {
    throw new Error(`Path escapes base directory: ${userPath}`);
  }
  return resolved;
}

module.exports = {
  assertSafeMediaPath,
  safeJoin,
};

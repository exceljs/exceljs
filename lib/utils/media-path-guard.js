'use strict';

const path = require('path');

// [exceljs-hardened] CVE fix — arbitrary local file read (GHSA pending).
//
// Workbook#addImage({filename}) / XLSX#addMedia() used to hand `filename`
// straight to fs.readFile() with zero validation. Any application that
// builds that path by joining a base directory with a user-supplied
// fragment (a very common "pick an asset by name" pattern) was exposed to
// classic ../ path traversal, letting a remote user read arbitrary files
// the Node process has access to (.env, source code, SSH keys, etc.).
//
// This guard rejects the traversal shape that made that exploitable:
// relative paths containing a ".." segment. It intentionally still allows
// absolute paths and plain relative filenames, since those are common,
// legitimate ways to reference a local file and aren't the vector that
// was exploited — the fix targets the traversal primitive itself, not
// the (correct, documented) ability to read a local file by path.
//
// This is defense in depth, not a substitute for validating untrusted
// input at the application layer: an application that forwards a raw,
// attacker-controlled *absolute* path is still responsible for its own
// allow-listing.
function assertSafeMediaPath(filename) {
  if (typeof filename !== 'string' || filename.length === 0) {
    return;
  }
  if (filename.includes('\0')) {
    throw new Error('Invalid media path: contains a null byte.');
  }
  const segments = path.normalize(filename).split(path.sep);
  if (segments.includes('..')) {
    throw new Error(
      `Invalid media path "${filename}": parent-directory traversal ("..") is not allowed. ` +
        'Pass an absolute path or a path with no ".." segments.'
    );
  }
}

// [exceljs-hardened] This is the fix that actually matters for the
// classic exploit chain — read it before relying on assertSafeMediaPath
// alone.
//
// The realistic vulnerable pattern is:
//   const resolvedPath = path.join(baseDir, userInput);
//   workbook.addImage({filename: resolvedPath});
//
// path.join() ALREADY resolves and strips ".." segments before that
// string ever reaches addImage() — e.g.
//   path.join('/app/assets/logos', '../../../.env') === '/app/.env'
// By the time assertSafeMediaPath() sees it, there is no ".." left to
// catch, and it (correctly, by its own contract) lets a fully-resolved
// absolute path through. No check performed *after* path.join() can
// undo that: the traversal information is already gone from the string.
//
// safeJoin() fixes this by doing the resolution and the containment
// check together, in the right order: resolve first, then verify the
// result is still inside baseDir. Applications should call this instead
// of path.join() wherever a user-influenced fragment is combined with a
// trusted base directory before being handed to addImage()/addMedia().
function safeJoin(baseDir, userPath) {
  const resolvedBase = path.resolve(baseDir);
  const resolvedTarget = path.resolve(resolvedBase, userPath);
  const relative = path.relative(resolvedBase, resolvedTarget);

  const escapesBase = relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative);
  if (escapesBase) {
    throw new Error(
      `Refusing to resolve "${userPath}" against base directory "${resolvedBase}": ` +
        `it resolves to "${resolvedTarget}", which is outside the base directory.`
    );
  }

  return resolvedTarget;
}

module.exports = {assertSafeMediaPath, safeJoin};

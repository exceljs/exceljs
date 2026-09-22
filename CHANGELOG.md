# Changelog

All notable changes to `exceljs-hardened` are documented here. This fork
starts from upstream `exceljs@4.4.0`; see the
[upstream changelog](https://github.com/exceljs/exceljs) for history
before that point.

## 5.0.0 — Security hardening release

First release of this fork. Patches all four vulnerabilities identified
in the audit of upstream `exceljs@4.4.0`, none of which could be
reported upstream — see [`SECURITY.md`](./SECURITY.md) for why.

Released as a major version bump, not a patch, because closing these
issues changes behavior for a few narrow but real inputs that used to be
accepted silently: see "Upgrading from 4.x" in the README before
updating.

### Fixed

- **Prototype pollution** in `_.deepMerge()` (`lib/utils/under-dash.js`).
  A JSON-parsed object containing an own `"__proto__"` key, merged via
  `Note.model` (e.g. through `cell.note = <attacker JSON>`), could reach
  and mutate the real global `Object.prototype`. `deepMerge()` now skips
  `__proto__`, `constructor`, and `prototype` keys unconditionally.
  Verified: `/api/whoami`-style checks built on a fresh plain object no
  longer flip after sending the payload.

- **Arbitrary local file read** via `addImage({filename})`
  (`lib/doc/workbook.js`, `lib/xlsx/xlsx.js`). The `filename` option was
  passed straight to `fs.readFile()` with no validation, so an
  application that built the path by joining a base directory with
  user input was exposed to classic `../` traversal.

  Two layers were needed to actually close this, in this order:
  1. `assertSafeMediaPath()` rejects any *raw* path containing a `..`
     segment handed directly to `addImage()`/`addMedia()`.
  2. `ExcelJS.utils.safeJoin(baseDir, userInput)` (new) — the fix that
     matters for the realistic exploit chain. A first draft of this
     patch shipped with only (1) and was re-tested against the original
     live PoC: it did nothing, because the demo app's own
     `path.join(baseDir, userInput)` already resolves and strips `..`
     *before* `addImage()` ever sees the string (7/9 test paths,
     including `.env`, were still exfiltrated). `safeJoin()` does
     resolution and containment-checking as a single operation instead
     of two separate steps, and is what application code needs to adopt
     in place of `path.join()`. Re-verified against the same PoC after
     the fix: 0/9 paths exfiltrated.

- **CSV / formula injection** in `CSV.write()` (`lib/csv/csv.js`).
  Leading `=`, `+`, `-`, `@`, tab, and CR characters were written to CSV
  output unescaped, letting a value that started with one of them be
  evaluated as a live formula by whoever opened the export in Excel or
  LibreOffice Calc. Such values are now prefixed with a neutralizing
  apostrophe before being handed to `fast-csv`, applied to both the
  default mapper and any caller-supplied `options.map`. Opt out per-call
  with `{escapeFormulas: false}` if needed.

- **Decompression bomb / memory exhaustion** in `Workbook#load()`
  (`lib/xlsx/xlsx.js`). Zip entries were fully decompressed into memory
  with no size cap, so a few-hundred-KB `.xlsx` could expand to
  gigabytes before any XML was parsed. `load()` now checks each entry's
  declared uncompressed size (read from the zip's central directory,
  before decompression) against a per-entry (128MB default) and
  archive-wide (512MB default) limit, both configurable via new
  `options.maxEntryUncompressedSize` / `options.maxTotalUncompressedSize`.
  Verified: a 1.5GB-decompressing payload is now rejected with an HTTP
  500 in the demo app instead of taking the process down (previously
  OOM-killed, exit 137, under a 512MB container limit).

### New public API

- `ExcelJS.utils.safeJoin(baseDir, userInput)`
- `Workbook#load()` options: `maxEntryUncompressedSize`, `maxTotalUncompressedSize`
- `CSV.write()`/`writeBuffer()`/`writeFile()` option: `escapeFormulas`

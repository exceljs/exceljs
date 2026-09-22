'use strict';

// [exceljs-hardened] CVE fix — CSV / formula injection (CWE-1236).
//
// CSV.write() used to hand cell values straight to fast-csv with no
// check for leading formula-trigger characters. A cell value beginning
// with =, +, -, or @ is interpreted as a live formula by Excel /
// LibreOffice Calc when the file is opened — enabling DDE-style command
// execution or data exfiltration (HYPERLINK/WEBSERVICE beacons) against
// whoever opens the export, even though the server never evaluated
// anything itself.
//
// The standard mitigation (OWASP's CSV injection guidance) is to prefix
// any such value with a leading apostrophe: every mainstream spreadsheet
// application treats a leading "'" as "force this cell to be text," and
// the apostrophe itself is not displayed to the user — the cell just
// renders as plain text instead of being evaluated.
const FORMULA_TRIGGER_CHARS = new Set(['=', '+', '-', '@', '\t', '\r']);

function escapeCsvFormulaValue(value) {
  if (typeof value !== 'string' || value.length === 0) {
    return value;
  }
  if (FORMULA_TRIGGER_CHARS.has(value[0])) {
    return `'${value}`;
  }
  return value;
}

module.exports = {escapeCsvFormulaValue, FORMULA_TRIGGER_CHARS};

# ExcelJS

This is a fork of the "exceljs" package, which fixes the problem with writing a file using streams and not using RAM.
It solves the problem of writing large exel files
<a href="https://github.com/exceljs/exceljs">Original repo</a>

# Installation

```shell
npm install @zlooun/exceljs
```

# Whats new!
To use streams corractly just write:

```javascript
import * as fs from 'fs';
import { stream } from 'exceljs';

const output_file_name = "/test.xlsx";

const writeStream = fs.createWriteStream(output_file_name, { flags: 'w' });
const wb = new stream.xlsx.WorkbookWriter({ stream: writeStream });
const worksheet = wb.addWorksheet("test");

const headers = Array.from({length: 256}, (_, i) => i + 1).map((i) => 'test' + i);

for (let i = 0; i < 100000; i++) {
  const row = headers.map((header) => header + '|' + i);
  await worksheet.addRow(row).commit(); // This raw will be immediately written to disk and will not clog RAM.
}

await worksheet.commit(); // This is not necessary because await wb.commit() is used, but you can also write to disk not raw by raw, but worksheet by worksheet.
await wb.commit();
```
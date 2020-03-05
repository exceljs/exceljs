const fs = require('fs');
const path = require('path');
const ExcelJS = require('./lib/exceljs.nodejs.js');

const runs = 3;

runProfiling('huge xlsx file', () => {
  return new Promise((resolve, reject) => {
    // Data taken from http://eforexcel.com/wp/downloads-18-sample-csv-files-data-sets-for-testing-sales/
    const stream = fs.createReadStream(path.join(__dirname, 'spec/integration/data/huge.xlsx'));

    const wb = new ExcelJS.stream.xlsx.WorkbookReader();
    wb.read(stream, {
      entries: 'emit',
      sharedStrings: 'cache',
      worksheets: 'emit',
    });

    let worksheetCount = 0;
    let rowCount = 0;
    wb.on('worksheet', worksheet => {
      worksheetCount += 1;
      console.log(`Reading worksheet ${worksheetCount}`);
      worksheet.on('row', () => {
        rowCount += 1;
        if (rowCount % 50000 === 0) console.log(`Reading row ${rowCount}`);
      });
    });

    wb.on('end', () => {
      console.log(`Processed ${worksheetCount} worksheets and ${rowCount} rows`);
      resolve();
    });
    wb.on('error', reject);
  });
});

async function runProfiling(name, run) {
  console.log('');
  console.log('####################################################');
  console.log(`WARMUP: Current memory usage: ${currentMemoryUsage({runGarbageCollector: true})} MB`);
  console.log(`WARMUP: ${name} profiling started`);
  const warmupStartTime = Date.now();
  await run();
  console.log(`WARMUP: ${name} profiling finished in ${Date.now() - warmupStartTime}ms`);
  console.log(`WARMUP: Current memory usage (before GC): ${currentMemoryUsage({runGarbageCollector: false})} MB`);
  console.log(`WARMUP: Current memory usage (after GC): ${currentMemoryUsage({runGarbageCollector: true})} MB`);

  for (let i = 1; i <= runs; i += 1) {
    console.log('');
    console.log('####################################################');
    console.log(`RUN ${i}: ${name} profiling started`);
    const startTime = Date.now();
    await run(); // eslint-disable-line no-await-in-loop
    console.log(`RUN ${i}: ${name} profiling finished in ${Date.now() - startTime}ms`);
    console.log(`RUN ${i}: Current memory usage (before GC): ${currentMemoryUsage({runGarbageCollector: false})} MB`);
    console.log(`RUN ${i}: Current memory usage (after GC): ${currentMemoryUsage({runGarbageCollector: true})} MB`);
  }
}

function currentMemoryUsage({runGarbageCollector}) {
  if (runGarbageCollector) global.gc();
  return Math.round((process.memoryUsage().heapUsed / 1024 / 1024) * 100) / 100;
}

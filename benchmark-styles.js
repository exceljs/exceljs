/* eslint-disable no-console, no-unused-vars */
const ExcelJS = require('./lib/exceljs.nodejs');

const runs = 3;
const N_ROWS = 1 * 200 * 1000;
const profileMap = {};

function generateRecords(n) {
  const records = [];
  for (let i = 0; i < n; i++) {
    records.push({
      id: '12313',
      myString1: 'asdasdas',
      myString2: 'asdasdasasdasdasdasdasdadsad',
      myNumericString: '234234',
      amount: 3425.34,
      myDate1: '2004-01-01',
      myDate2: '2005-01-01',
    });
  }
  return records;
}

async function writeToExcel(records, stylesCacheMode) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: `benchmark-styles-${stylesCacheMode}.xlsx`,
    useStyles: stylesCacheMode !== 'NO_STYLES',
    stylesCacheMode,
  });

  const columns = [
    {header: 'ID', key: 'id', width: 22},
    {header: 'My String 1', key: 'myString1', width: 22},
    {header: 'My Numeric String', key: 'myNumericString', width: 22},
    {header: 'My Strign 2', key: 'myString2', width: 22},
    {header: 'Amount', key: 'amount', width: 15, style: {numFmt: '0.0'}},
    {header: 'My Date 1', key: 'myDate1', width: 15, style: {numFmt: 'yyyy-mm-dd'}},
    {header: 'My Date 2', key: 'myDate2', width: 15, style: {numFmt: 'yyyy-mm-dd'}},
  ];

  await createSheet(workbook, 1, columns, records, stylesCacheMode);

  await workbook.commit();
}

async function createSheet(workbook, i, columns, records) {
  const worksheet = workbook.addWorksheet(`Sheet${i}`);
  worksheet.columns = columns;

  for (const record of records) {
    const row = worksheet.addRow(record);
    row.eachCell(cell => {
      cell.style = {
        border: {top: {style: 'thin'}, left: {style: 'thin'}, bottom: {style: 'thin'}, right: {style: 'thin'}},
        font: {name: 'Times New Roman', size: 10},
      };
    });
    row.commit();
  }

  await worksheet.commit();
}

(async () => {
  try {
    const records = generateRecords(N_ROWS);
    await runProfiling('NO_STYLES', async () => {
      await writeToExcel(records, 'NO_STYLES');
    });
    await runProfiling('WEAK_MAP', async () => {
      await writeToExcel(records);
    });
    await runProfiling('JSON_MAP', async () => {
      await writeToExcel(records, 'JSON_MAP');
    });
    await runProfiling('FAST_MAP', async () => {
      await writeToExcel(records, 'FAST_MAP');
    });
    await runProfiling('NO_CACHE', async () => {
      await writeToExcel(records, 'NO_CACHE');
    });
    const baseline = profileMap.NO_STYLES.reduce((a, v) => a + v, 0) / runs / 1000;
    console.log('\n\nMode\t\tAVG (s)\t\tx NO_STYLES');
    Object.keys(profileMap).forEach(key => {
      const prof = profileMap[key];
      const average = prof.reduce((a, v) => a + v, 0) / runs / 1000;
      const x = average / baseline;
      console.log(`${key}\t\t${average.toFixed(2)}\t\t${x.toFixed(2)}`);
    });
    // console.log(profileMap)
  } catch (err) {
    console.error(err);
    process.exit(1); // eslint-disable-line no-process-exit
  }
})();

async function runProfiling(name, run) {
  console.log('');
  console.log('####################################################');
  console.log(`WARMUP: Current memory usage: ${currentMemoryUsage({runGarbageCollector: true})} MB`);
  console.log(`WARMUP: ${name} profiling started`);
  const warmupStartTime = Date.now();
  await run();
  console.log(`WARMUP: ${name} profiling finished in ${Date.now() - warmupStartTime}ms`);
  console.log(
    `WARMUP: Current memory usage (before GC): ${currentMemoryUsage({
      runGarbageCollector: false,
    })} MB`
  );
  console.log(`WARMUP: Current memory usage (after GC): ${currentMemoryUsage({runGarbageCollector: true})} MB`);

  const prof = profileMap[name] || [];
  profileMap[name] = prof;
  for (let i = 1; i <= runs; i += 1) {
    console.log('');
    console.log('####################################################');
    console.log(`RUN ${i}: ${name} profiling started`);
    const startTime = Date.now();
    await run(); // eslint-disable-line no-await-in-loop
    const duration = Date.now() - startTime;
    prof.push(duration);
    console.log(`RUN ${i}: ${name} profiling finished in ${duration}ms`);

    console.log(
      `RUN ${i}: Current memory usage (before GC): ${currentMemoryUsage({
        runGarbageCollector: false,
      })} MB`
    );
    console.log(
      `RUN ${i}: Current memory usage (after GC): ${currentMemoryUsage({
        runGarbageCollector: true,
      })} MB`
    );
  }
}

function currentMemoryUsage({runGarbageCollector}) {
  if (runGarbageCollector) global.gc();
  return Math.round((process.memoryUsage().heapUsed / 1024 / 1024) * 100) / 100;
}

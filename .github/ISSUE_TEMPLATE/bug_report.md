---
name: Bug report
about: Create a report to help us improve
title: "[BUG] XYZ"
labels: bug
assignees: ''

---

<!--
Bug description + screenshots
-->

Lib version:

## Steps To Reproduce

1.
2.
3.

### code example:

```javascript
/* please paste code here */

```

### test code (optionally, helpful):

```javascript
/* if you have or could write one It should help us quickly a lot :) */ 

/* example test code: */

const path = require('path');
const ExcelJS = require('exceljs');

describe('Bug issue test', () => {
    it('issue #xyz test - TOPIC', () => { /* it help us in future grab more information when test will faild again */

      // prepare
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('XYZ');

      // action
      ws.getCell('A1').value = 7;
       /* ... */      


      // checks
      expect(ws.getCell('A1').value).to.equal(7);
    });
});

```


## The current behaviour:


## The expected behaviour:

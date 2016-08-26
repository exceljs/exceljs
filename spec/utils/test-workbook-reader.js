'use strict';

var expect = require('chai').expect;
var Bluebird = require('bluebird');
var _ = require('underscore');
var tools = require('./tools');

module.exports = {
  testValues: tools.fix(require('./data/sheet-values.json')),
  styles: tools.fix(require('./data/styles.json')),
  properties: tools.fix(require('./data/sheet-properties.json')),
  pageSetup: tools.fix(require('./data/page-setup.json')),

  checkBook: function(filename) {
    var wb = new Excel.stream.xlsx.WorkbookReader();

    // expectations
    var dateAccuracy = 3;

    var deferred = Bluebird.defer();

    wb.on('worksheet', function(ws) {
      // Sheet name stored in workbook. Not guaranteed here
      // expect(ws.name).to.equal('blort');
      ws.on('row', function(row) {
        switch(row.number) {
          case 1:
            expect(row.getCell('A').value).to.equal(7);
            expect(row.getCell('A').type).to.equal(Excel.ValueType.Number);
            expect(row.getCell('B').value).to.equal(utils.testValues.str);
            expect(row.getCell('B').type).to.equal(Excel.ValueType.String);
            expect(Math.abs(row.getCell('C').value.getTime() - utils.testValues.date.getTime())).to.be.below(dateAccuracy);
            expect(row.getCell('C').type).to.equal(Excel.ValueType.Date);

            expect(row.getCell('D').value).to.deep.equal(utils.testValues.formulas[0]);
            expect(row.getCell('D').type).to.equal(Excel.ValueType.Formula);
            expect(row.getCell('E').value).to.deep.equal(utils.testValues.formulas[1]);
            expect(row.getCell('E').type).to.equal(Excel.ValueType.Formula);
            expect(row.getCell('F').value).to.deep.equal(utils.testValues.hyperlink);
            expect(row.getCell('F').type).to.equal(Excel.ValueType.Hyperlink);
            expect(row.getCell('G').value).to.equal(utils.testValues.str2);
            break;

          case 2:
            // A2:B3
            expect(row.getCell('A').value).to.equal(5);
            expect(row.getCell('A').type).to.equal(Excel.ValueType.Number);
            //expect(row.getCell('A').master).to.equal(ws.getCell('A2'));

            expect(row.getCell('B').value).to.equal(5);
            expect(row.getCell('B').type).to.equal(Excel.ValueType.Merge);
            //expect(row.getCell('B').master).to.equal(ws.getCell('A2'));

            // C2:D3
            expect(row.getCell('C').value).to.be.null;
            expect(row.getCell('C').type).to.equal(Excel.ValueType.Null);
            //expect(row.getCell('C').master).to.equal(ws.getCell('C2'));

            expect(row.getCell('D').value).to.be.null;
            expect(row.getCell('D').type).to.equal(Excel.ValueType.Merge);
            //expect(row.getCell('D').master).to.equal(ws.getCell('C2'));

            break;

          case 3:
            expect(row.getCell('A').value).to.equal(5);
            expect(row.getCell('A').type).to.equal(Excel.ValueType.Merge);
            //expect(row.getCell('A').master).to.equal(ws.getCell('A2'));

            expect(row.getCell('B').value).to.equal(5);
            expect(row.getCell('B').type).to.equal(Excel.ValueType.Merge);
            //expect(row.getCell('B').master).to.equal(ws.getCell('A2'));

            expect(row.getCell('C').value).to.be.null;
            expect(row.getCell('C').type).to.equal(Excel.ValueType.Merge);
            //expect(row.getCell('C').master).to.equal(ws.getCell('C2'));

            expect(row.getCell('D').value).to.be.null;
            expect(row.getCell('D').type).to.equal(Excel.ValueType.Merge);
            //expect(row.getCell('D').master).to.equal(ws.getCell('C2'));
            break;

          case 4:
            expect(row.getCell('A').numFmt).to.equal(utils.testValues.numFmt1);
            expect(row.getCell('A').type).to.equal(Excel.ValueType.Number);
            expect(row.getCell('A').border).to.deep.equal(utils.styles.borders.thin);
            expect(row.getCell('C').numFmt).to.equal(utils.testValues.numFmt2);
            expect(row.getCell('C').type).to.equal(Excel.ValueType.Number);
            expect(row.getCell('C').border).to.deep.equal(utils.styles.borders.doubleRed);
            expect(row.getCell('E').border).to.deep.equal(utils.styles.borders.thickRainbow);
            break;

          case 5:
            // test fonts and formats
            expect(row.getCell('A').value).to.equal(utils.testValues.str);
            expect(row.getCell('A').type).to.equal(Excel.ValueType.String);
            expect(row.getCell('A').font).to.deep.equal(utils.styles.fonts.arialBlackUI14);
            expect(row.getCell('B').value).to.equal(utils.testValues.str);
            expect(row.getCell('B').type).to.equal(Excel.ValueType.String);
            expect(row.getCell('B').font).to.deep.equal(utils.styles.fonts.broadwayRedOutline20);
            expect(row.getCell('C').value).to.equal(utils.testValues.str);
            expect(row.getCell('C').type).to.equal(Excel.ValueType.String);
            expect(row.getCell('C').font).to.deep.equal(utils.styles.fonts.comicSansUdB16);

            expect(Math.abs(row.getCell('D').value - 1.6)).to.be.below(0.00000001);
            expect(row.getCell('D').type).to.equal(Excel.ValueType.Number);
            expect(row.getCell('D').numFmt).to.equal(utils.testValues.numFmt1);
            expect(row.getCell('D').font).to.deep.equal(utils.styles.fonts.arialBlackUI14);

            expect(Math.abs(row.getCell('E').value - 1.6)).to.be.below(0.00000001);
            expect(row.getCell('E').type).to.equal(Excel.ValueType.Number);
            expect(row.getCell('E').numFmt).to.equal(utils.testValues.numFmt2);
            expect(row.getCell('E').font).to.deep.equal(utils.styles.fonts.broadwayRedOutline20);

            expect(Math.abs(ws.getCell('F5').value.getTime() - utils.testValues.date.getTime())).to.be.below(dateAccuracy);
            expect(ws.getCell('F5').type).to.equal(Excel.ValueType.Date);
            expect(ws.getCell('F5').numFmt).to.equal(utils.testValues.numFmtDate);
            expect(ws.getCell('F5').font).to.deep.equal(utils.styles.fonts.comicSansUdB16);
            expect(row.height).to.be.undefined;
            break;

          case 6:
            expect(ws.getRow(6).height).to.equal(42);
            _.each(utils.styles.alignments, function(alignment, index) {
              var colNumber = index + 1;
              var cell = row.getCell(colNumber);
              expect(cell.value).to.equal(alignment.text);
              expect(cell.alignment).to.deep.equal(alignment.alignment);
            });
            break;

          case 7:
            _.each(utils.styles.badAlignments, function(alignment, index) {
              var colNumber = index + 1;
              var cell = row.getCell(colNumber);
              expect(cell.value).to.equal(alignment.text);
              expect(cell.alignment).to.be.undefined;
            });
            break;

          case 8:
            expect(row.height).to.equal(40);
            expect(row.getCell(1).fill).to.deep.equal(utils.styles.fills.blueWhiteHGrad);
            expect(row.getCell(2).fill).to.deep.equal(utils.styles.fills.redDarkVertical);
            expect(row.getCell(3).fill).to.deep.equal(utils.styles.fills.redGreenDarkTrellis);
            expect(row.getCell(4).fill).to.deep.equal(utils.styles.fills.rgbPathGrad);
            break;
        }
      });
    });
    wb.on('end', function() {
      deferred.resolve();
    });

    wb.read(filename, {entries: 'emit', worksheets: 'emit'});

    return deferred.promise;
  }
  
};
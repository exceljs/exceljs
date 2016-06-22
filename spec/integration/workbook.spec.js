'use strict';

var stream = require('stream');
var expect = require('chai').expect;
var bluebird = require('bluebird');
var fs = require('fs');
var fsa = bluebird.promisifyAll(fs);
var _ = require('underscore');
var Excel = require('../../excel');
var testUtils = require('./../testutils');

// =============================================================================
// Sample Data
var richTextSample = require('./data/rich-text-sample');
var richTextSample_A1 = require('./data/rich-text-sample-a1.json');

// =============================================================================
// Tests

describe('Workbook', function() {

  describe('Serialise', function() {

    after(function() {
      function deleteFile(filename) {
        return fsa.unlinkAsync(filename)
          .catch(function() { });
      }

      return bluebird.all([
        deleteFile('./wb.test.xlsx'),
        deleteFile('./wb.test.csv')
      ]);
    });

    it('creates sheets and saves with correct names', function() {
      var wb = new Excel.Workbook();
      var ws1 = wb.addWorksheet('Hello, World!');
      expect(ws1.name).to.equal('Hello, World!');

      var ws2 = wb.addWorksheet();
      expect(ws2.name).to.match(/sheet\d+/);

      var ws3 = wb.addWorksheet('This & That');

      return wb.xlsx.writeFile('./wb.test.xlsx')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wb.test.xlsx');
        })
        .then(function(wb2) {
          expect(wb2.getWorksheet('Hello, World!')).to.be.ok;
          expect(wb2.getWorksheet('This & That')).to.be.ok;
        });
    });

    it('serializes and deserializes to xlsx file properly', function() {

      var wb = testUtils.createTestBook(true, Excel.Workbook);
      //fs.writeFileSync('./testmodel.json', JSON.stringify(wb.model, null, '    '));

      return wb.xlsx.writeFile('./wb.test.xlsx')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wb.test.xlsx');
        })
        .then(function(wb2) {
          testUtils.checkTestBook(wb2, 'xlsx', true);
        });
    });

    it('serializes and deserializes dataValidations', function() {
      var wb = new Excel.Workbook();
      testUtils.addDataValidationSheet(wb);

      return wb.xlsx.writeFile('./wb.test.xlsx')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wb.test.xlsx');
        })
        .then(function(wb2) {
          testUtils.checkDataValidationSheet(wb2);
        });
    });

    it("Adds an empty string", function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet();

      ws.columns = [
        { key: "id", width: 10 },
        { key: "name", width: 32 }
      ];

      ws.addRow({id: 1, name: ''});

      return wb.xlsx.writeFile('./wb.test.xlsx');
    });

    it('serializes row styles and columns properly', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.columns = [
        { header: 'A1', width: 10 },
        { header: 'B1', width: 20, style: { font: testUtils.styles.fonts.comicSansUdB16, alignment: testUtils.styles.alignments[1].alignment } },
        { header: 'C1', width: 30 }
      ];

      ws.getRow(2).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('C2').value = 'C2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';
      ws.getCell('C3').value = 'C3';

      return wb.xlsx.writeFile('./wb.test.xlsx')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wb.test.xlsx');
        })
        .then(function(wb2) {
          var ws2 = wb2.getWorksheet('blort');
          _.each(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'], function(address) {
            expect(ws2.getCell(address).value).to.equal(address);
          });
          expect(ws2.getCell('B1').font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getCell('B1').alignment).to.deep.equal(testUtils.styles.alignments[1].alignment);
          expect(ws2.getCell('A2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('B2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('C2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('B3').font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getCell('B3').alignment).to.deep.equal(testUtils.styles.alignments[1].alignment);

          expect(ws2.getColumn(2).font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getColumn(2).alignment).to.deep.equal(testUtils.styles.alignments[1].alignment);

          expect(ws2.getRow(2).font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
        });
    });

    it('serializes and deserializes a lot of sheets to xlsx file properly', function() {
      this.timeout(10000);

      var i;
      var wb = new Excel.Workbook();
      var numSheets = 90;
      // add numSheets sheets
      for (i = 1; i <= numSheets; i++) {
        var ws = wb.addWorksheet('sheet' + i);
        ws.getCell('A1').value = i;
      }
      return wb.xlsx.writeFile('./wb.test.xlsx')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wb.test.xlsx');
        })
        .then(function(wb2) {
          for (i = 1; i <= numSheets; i++) {
            var ws2 = wb2.getWorksheet('sheet' + i);
            expect(ws2).to.be.ok;
            expect(ws2.getCell('A1').value).to.equal(i);
          }
        });
    });

    it('deserializes in-cell formats properly in xlsx file', function() {

      // Stream from input string
      var testData = new Buffer(richTextSample, 'base64');

      // Initiate the source
      var bufferStream = new stream.PassThrough();

      // Write your buffer
      bufferStream.write(testData);
      bufferStream.end();

      var wb = new Excel.Workbook();
      return wb.xlsx.read(bufferStream)
          .then(function () {
            var ws = wb.worksheets[0];
            expect(ws.getCell("A1").value).to.deep.equal(richTextSample_A1);
            expect(ws.getCell("A1").text).to.equal(ws.getCell("A2").value);
          });
    });

    it('serializes and deserialises to csv file properly', function() {
      this.timeout(5000);

      var wb = testUtils.createTestBook(true, Excel.Workbook);
      //fs.writeFileSync('./testmodel.json', JSON.stringify(wb.model, null, '    '));

      return wb.csv.writeFile('./wb.test.csv')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.csv.readFile('./wb.test.csv')
            .then(function() {
              return wb2;
            });
        })
        .then(function(wb2) {
          testUtils.checkTestBook(wb2, 'csv');
        });
    });

    it('serialises and deserialises defined names', function() {
      var wb1 = new Excel.Workbook();
      var ws1a = wb1.addWorksheet('blort');
      var ws1b = wb1.addWorksheet('foo');

      function assign(sheet, address, value, name) {
        var cell = sheet.getCell(address);
        cell.value = value;
        if (name instanceof  Array) {
          cell.names = name;
        } else {
          cell.name = name;
        }
      }

      // single entry
      assign(ws1a, 'A1', 5, 'five');

      // three amigos - horizontal line
      assign(ws1a, 'A3', 3, 'amigos');
      assign(ws1a, 'B3', 3, 'amigos');
      assign(ws1a, 'C3', 3, 'amigos');

      // three amigos - vertical line
      assign(ws1a, 'E1', 3, 'verts');
      assign(ws1a, 'E2', 3, 'verts');
      assign(ws1a, 'E3', 3, 'verts');

      // four square
      assign(ws1a, 'C5', 4, 'squares');
      assign(ws1a, 'B6', 4, 'squares');
      assign(ws1a, 'C6', 4, 'squares');
      assign(ws1a, 'B5', 4, 'squares');

      // long distance
      assign(ws1a, 'B7', 2, 'sheets');
      assign(ws1b, 'B7', 2, 'sheets');

      // two names
      assign(ws1a, 'G1', 1, 'thing1');
      ws1a.getCell('G1').addName('thing2');

      // once removed
      assign(ws1a, 'G2', 1, ['once', 'twice']);
      ws1a.getCell('G2').removeName('once');

      return wb1.xlsx.writeFile('./wb.test.xlsx')
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wb.test.xlsx');
        })
        .then(function(wb2) {
          var ws2a = wb2.getWorksheet('blort');
          var ws2b = wb2.getWorksheet('foo');

          function check(sheet, address, value, name) {
            var cell = sheet.getCell(address);
            expect(cell.value).to.equal(value);
            expect(cell.name).to.equal(name);
          }

          // single entry
          check(ws2a, 'A1', 5, 'five');

          // three amigos - horizontal line
          check(ws2a, 'A3', 3, 'amigos');
          check(ws2a, 'B3', 3, 'amigos');
          check(ws2a, 'C3', 3, 'amigos');

          // three amigos - vertical line
          check(ws2a, 'E1', 3, 'verts');
          check(ws2a, 'E2', 3, 'verts');
          check(ws2a, 'E3', 3, 'verts');

          // four square
          check(ws2a, 'C5', 4, 'squares');
          check(ws2a, 'B6', 4, 'squares');
          check(ws2a, 'C6', 4, 'squares');
          check(ws2a, 'B5', 4, 'squares');

          // long distance
          check(ws2a, 'B7', 2, 'sheets');
          check(ws2b, 'B7', 2, 'sheets');

          // two names
          expect(ws2a.getCell('G1').names).to.have.members(['thing1', 'thing2']);

          // once removed
          expect(ws2a.getCell('G2').names).to.have.members(['twice']);

          // ranges
          function rangeCheck(name, members) {
            var ranges = wb2.definedNames.getRanges(name);
            expect(ranges.name).to.equal(name);
            if (members.length) {
              expect(ranges.ranges).to.have.members(members);
            } else {
              expect(ranges.ranges.length).to.equal(0);
            }
          }
          rangeCheck('five', ['blort!$A$1']);
          rangeCheck('amigos', ['blort!$A$3:$C$3']);
          rangeCheck('verts', ['blort!$E$1:$E$3']);
          rangeCheck('squares', ['blort!$B$5:$C$6']);
          rangeCheck('sheets', ['blort!$B$7','foo!$B$7']);
          rangeCheck('thing1', ['blort!$G$1']);
          rangeCheck('thing2', ['blort!$G$1']);
          rangeCheck('once', []);
          rangeCheck('twice', ['blort!$G$2']);
        });
    });

    describe('Merge Cells', function() {
      it('serialises and deserialises properly', function() {
        var wb = new Excel.Workbook();
        var ws = wb.addWorksheet('blort');

        // initial values
        ws.getCell('B2').value = 'B2';

        ws.mergeCells('B2:C3');

        return wb.xlsx.writeFile('./wb.test.xlsx')
          .then(function() {
            var wb2 = new Excel.Workbook();
            return wb2.xlsx.readFile('./wb.test.xlsx');
          })
          .then(function(wb2) {
            var ws2 = wb2.getWorksheet('blort');

            expect(ws2.getCell('B2').value).to.equal('B2');
            expect(ws2.getCell('B3').value).to.equal('B2');
            expect(ws2.getCell('C2').value).to.equal('B2');
            expect(ws2.getCell('C3').value).to.equal('B2');

            expect(ws2.getCell('B2').type).to.equal(Excel.ValueType.String);
            expect(ws2.getCell('B3').type).to.equal(Excel.ValueType.Merge);
            expect(ws2.getCell('C2').type).to.equal(Excel.ValueType.Merge);
            expect(ws2.getCell('C3').type).to.equal(Excel.ValueType.Merge);
          });
      });

      it('serialises and deserialises styles', function() {
        var wb = new Excel.Workbook();
        var ws = wb.addWorksheet('blort');

        // initial values
        var B2 = ws.getCell('B2');
        B2.value = 5;
        B2.style.font = testUtils.styles.fonts.broadwayRedOutline20;
        B2.style.border = testUtils.styles.borders.doubleRed;
        B2.style.fill = testUtils.styles.fills.blueWhiteHGrad;
        B2.style.alignment = testUtils.styles.namedAlignments.middleCentre;
        B2.style.numFmt = testUtils.styles.numFmts.numFmt1;

        // expecting styles to be copied (see worksheet spec)
        ws.mergeCells('B2:C3');

        return wb.xlsx.writeFile('./wb.test.xlsx')
          .then(function() {
            var wb2 = new Excel.Workbook();
            return wb2.xlsx.readFile('./wb.test.xlsx');
          })
          .then(function(wb2) {
            var ws2 = wb2.getWorksheet('blort');

            expect(ws2.getCell('B2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('B2').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('B2').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('B2').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('B2').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);

            expect(ws2.getCell('B3').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('B3').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('B3').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('B3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('B3').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);

            expect(ws2.getCell('C2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('C2').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('C2').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('C2').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('C2').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);

            expect(ws2.getCell('C3').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('C3').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('C3').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('C3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('C3').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);
          });
      });
    });
  });

  it('serialises and deserialises by model', function() {
    var wb = testUtils.createTestBook(false, Excel.Workbook);

    return testUtils.cloneByModel(wb, Excel.Workbook)
      .then(function(wb2) {
        testUtils.checkTestBook(wb2, 'model');
      });
  });

  it('throws an error when xlsx file not found', function() {
    var wb = new Excel.Workbook();
    var success = 0;
    return wb.xlsx.readFile('./wb.doesnotexist.xlsx')
      .then(function(wb) {
        success = 1;
      })
      .catch(function(error) {
        success = 2;
        // expect the right kind of error
      })
      .finally(function() {
        expect(success).to.equal(2);
      });
  });

  it('throws an error when csv file not found', function() {
    var wb = new Excel.Workbook();
    var success = 0;
    return wb.csv.readFile('./wb.doesnotexist.csv')
      .then(function(wb) {
        success = 1;
      })
      .catch(function(error) {
        success = 2;
        // expect the right kind of error
      })
      .finally(function() {
        expect(success).to.equal(2);
      });
  });
});

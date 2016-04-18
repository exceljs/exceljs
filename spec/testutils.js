'use strict';

var expect = require('chai').expect;

var Bluebird = require('bluebird');
var _ = require('underscore');
var MemoryStream = require('memorystream');

var Excel = require('../excel');



var utils = module.exports = {
  concatenateFormula: function() {
    var args = Array.prototype.slice.call(arguments);
    var values = args.map(function(value) {
      return '"' + value + '"';
    });
    return {
      formula: 'CONCATENATE(' + values.join(',') + ')'
    };
  },
  cloneByModel: function(thing1, Type) {
    var model = thing1.model;
    //console.log(JSON.stringify(model, null, '    '))
    var thing2 = new Type();
    thing2.model = model;
    return Bluebird.resolve(thing2);
  },
  cloneByStream: function(thing1, Type, end) {
    var deferred = Bluebird.defer();
    end = end || 'end';

    var thing2 = new Type();
    var stream = thing2.createInputStream();
    stream.on(end, function() {
      deferred.resolve(thing2);
    });
    stream.on('error', function(error) {
      deferred.reject(error);
    });

    var memStream = new MemoryStream();
    memStream.on('error', function(error) {
      deferred.reject(error);
    });
    memStream.pipe(stream);
    thing1.write(memStream)
      .then(function() {
        memStream.end();
      });

    return deferred.promise;
  },
  testValues: {
    num: 7,
    str: 'Hello, World!',
    str2: '<a href="www.whatever.com">Talk to the H&</a>',
    date: new Date(),
    formulas: [
      {formula: 'A1', result: 7},
      {formula: 'A2', result: undefined}
    ],
    hyperlink: {hyperlink: 'http://www.link.com', text: 'www.link.com'},
    numFmt1: '# ?/?',
    numFmt2: '[Green]#,##0 ;[Red](#,##0)',
    numFmtDate: 'dd, mmm yyyy'
  },
  styles: {
    numFmts: {
      numFmt1: '# ?/?',
      numFmt2: '[Green]#,##0 ;[Red](#,##0)'
    },
    fonts: {
      arialBlackUI14: { name: 'Arial Black', family: 2, size: 14, underline: true, italic: true },
      comicSansUdB16: { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true },
      broadwayRedOutline20: { name: 'Broadway', family: 5, size: 20, outline: true, color: { argb:'FFFF0000'}}
    },
    alignments: [
      { text: 'Top Left', alignment: { horizontal: 'left', vertical: 'top' } },
      { text: 'Middle Centre', alignment: { horizontal: 'center', vertical: 'middle' } },
      { text: 'Bottom Right', alignment: { horizontal: 'right', vertical: 'bottom' } },
      { text: 'Wrap Text', alignment: { wrapText: true } },
      { text: 'Indent 1', alignment: { indent: 1 } },
      { text: 'Indent 2', alignment: { indent: 2 } },
      { text: 'Rotate 15', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 15 } },
      { text: 'Rotate 30', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 30 } },
      { text: 'Rotate 45', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 45 } },
      { text: 'Rotate 60', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 60 } },
      { text: 'Rotate 75', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 75 } },
      { text: 'Rotate 90', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 90 } },
      { text: 'Rotate -15', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: -55 } },
      { text: 'Rotate -30', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: -30 } },
      { text: 'Rotate -45', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: -45 } },
      { text: 'Rotate -60', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: -60 } },
      { text: 'Rotate -75', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: -75 } },
      { text: 'Rotate -90', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: -90 } },
      { text: 'Vertical Text', alignment: { horizontal: 'right', vertical: 'bottom', textRotation: 'vertical' } }
    ],
    namedAlignments: {
      topLeft: { horizontal: 'left', vertical: 'top' },
      middleCentre: { horizontal: 'center', vertical: 'middle' },
      bottomRight: { horizontal: 'right', vertical: 'bottom' }
    },
    badAlignments: [
      { text: 'Rotate -91', alignment: { textRotation: -91 } },
      { text: 'Rotate 91', alignment: { textRotation: 91 } },
      { text: 'Indent -1', alignment: { indent: -1 } },
      { text: 'Blank', alignment: {  } }
    ],
    borders: {
      thin: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}},
      doubleRed: { top: {style:'double', color: {argb:'FFFF0000'}}, left: {style:'double', color: {argb:'FFFF0000'}}, bottom: {style:'double', color: {argb:'FFFF0000'}}, right: {style:'double', color: {argb:'FFFF0000'}}},
      thickRainbow: {
        top: {style:'double', color: {argb:'FFFF00FF'}},
        left: {style:'double', color: {argb:'FF00FFFF'}},
        bottom: {style:'double', color: {argb:'FF00FF00'}},
        right: {style:'double', color: {argb:'FF00FF'}},
        diagonal: {style:'double', color: {argb:'FFFFFF00'}, up: true, down: true}
      }
    },
    fills: {
      redDarkVertical: {type: 'pattern', pattern:'darkVertical', fgColor:{argb:'FFFF0000'}},
      redGreenDarkTrellis: {type: 'pattern', pattern:'darkTrellis',
        fgColor:{argb:'FFFF0000'}, bgColor:{argb:'FF00FF00'}},
      blueWhiteHGrad: {type: 'gradient', gradient: 'angle', degree: 0,
        stops: [{position:0, color:{argb:'FF0000FF'}},{position:1, color:{argb:'FFFFFFFF'}}]},
      rgbPathGrad: {type: 'gradient', gradient: 'path', center:{left:0.5,top:0.5},
        stops: [
          {position:0, color:{argb:'FFFF0000'}},
          {position:0.5, color:{argb:'FF00FF00'}},
          {position:1, color:{argb:'FF0000FF'}}
        ]
      }
    }
  },

  dataValidations: {
    B1: {
      type: 'list',
      allowBlank: true,
      formulae: ['Nephews']
    },

    B3: {
      type: 'list',
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: ['"One,Two,Three,Four"']
    },

    B5: {
      type: 'list',
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: ['$D$5:$F$5']
    },

    B13: {
      type: 'whole',
      operator: 'equal',
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: [5],
      promptTitle: 'Five',
      prompt: 'The value must be Five'
    },

    E13: {
      type: 'whole',
      operator: 'notEqual',
      allowBlank: true,
      showInputMessage: true,
      showErrorMessage: true,
      formulae: [5],
      errorStyle: 'error',
      errorTitle: 'Five',
      error: 'The value must not be Five'
    },

    B15: {
      type: 'whole',
      operator: 'notEqual',
      formulae: [5]
    },

    types: ['whole', 'decimal', 'date', 'textLength'],
    values: {
      whole: { v1: 1, v2: 10 },
      decimal: { v1: 1.5, v2: 10.2 },
      date: { v1: new Date(2015, 0, 1), v2: new Date(2016, 0, 1) },
      textLength: { v1: 5, v2: 15 }
    },
    operators: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
    create: function(type, operator) {
      var dataValidation = {
        type: type,
        operator: operator,
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: [this.values[type].v1]
      };
      switch(operator) {
        case 'between':
        case 'notBetween':
          dataValidation.formulae.push(this.values[type].v2);
          break;
      }
      return dataValidation;
    }
  },

  addDataValidationSheet: function(wb) {
    var ws = wb.addWorksheet('data-validations');

    // named list
    ws.getCell('D1').value = 'Hewie';
    ws.getCell('D1').name = 'Nephews';
    ws.getCell('E1').value = 'Dewie';
    ws.getCell('E1').name = 'Nephews';
    ws.getCell('F1').value = 'Louie';
    ws.getCell('F1').name = 'Nephews';
    ws.getCell('A1').value = utils.concatenateFormula('Named List');
    ws.getCell('B1').dataValidation = this.dataValidations.B1;

    ws.getCell('A3').value = utils.concatenateFormula('Literal List');
    ws.getCell('B3').dataValidation = this.dataValidations.B3;

    ws.getCell('D5').value = 'Tom';
    ws.getCell('E5').value = 'Dick';
    ws.getCell('F5').value = 'Harry';
    ws.getCell('A5').value = utils.concatenateFormula('Range List');
    ws.getCell('B5').dataValidation = this.dataValidations.B5;

    _.each(utils.dataValidations.operators, function(operator, cIndex) {
      var col = 3 + cIndex;
      ws.getCell(7, col).value = utils.concatenateFormula(operator);
    });
    _.each(utils.dataValidations.types, function(type, rIndex) {
      var row = 8 + rIndex;
      ws.getCell(row, 1).value = utils.concatenateFormula(type);
      _.each(utils.dataValidations.operators, function(operator, cIndex) {
        var col = 3 + cIndex;
        ws.getCell(row, col).dataValidation = utils.dataValidations.create(type, operator);
      });
    });

    ws.getCell('A13').value = utils.concatenateFormula('Prompt');
    ws.getCell('B13').dataValidation = this.dataValidations.B13;

    ws.getCell('D13').value = utils.concatenateFormula('Error');
    ws.getCell('E13').dataValidation = this.dataValidations.E13;

    ws.getCell('A15').value = utils.concatenateFormula('Terse');
    ws.getCell('B15').dataValidation = this.dataValidations.B15;

  },
  
  checkDataValidationSheet: function(wb) {
    var ws = wb.getWorksheet('data-validations');
    expect(ws).to.not.be.undefined;

    expect(ws.getCell('B1').dataValidation).to.deep.equal(this.dataValidations.B1);
    expect(ws.getCell('B3').dataValidation).to.deep.equal(this.dataValidations.B3);
    expect(ws.getCell('B5').dataValidation).to.deep.equal(this.dataValidations.B5);

    _.each(utils.dataValidations.types, function(type, rIndex) {
      var row = 8 + rIndex;
      ws.getCell(row, 1).value = utils.concatenateFormula(type);
      _.each(utils.dataValidations.operators, function(operator, cIndex) {
        var col = 3 + cIndex;
        expect(ws.getCell(row, col).dataValidation).to.deep.equal(utils.dataValidations.create(type, operator));
      });
    });

    expect(ws.getCell('B13').dataValidation).to.deep.equal(this.dataValidations.B13);
    expect(ws.getCell('E13').dataValidation).to.deep.equal(this.dataValidations.E13);
    expect(ws.getCell('B15').dataValidation).to.deep.equal(this.dataValidations.B15);
  },

  createTestBook: function(checkBadAlignments, WorkbookClass, options) {
    var wb = new WorkbookClass(options);
    var ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 7;
    ws.getCell('B1').value = utils.testValues.str;
    ws.getCell('C1').value = utils.testValues.date;
    ws.getCell('D1').value = utils.testValues.formulas[0];
    ws.getCell('E1').value = utils.testValues.formulas[1];
    ws.getCell('F1').value = utils.testValues.hyperlink;
    ws.getCell('G1').value = utils.testValues.str2;
    ws.getRow(1).commit();

    // merge cell square with numerical value
    ws.getCell('A2').value = 5;
    ws.mergeCells('A2:B3');

    // merge cell square with null value
    ws.mergeCells('C2:D3');
    ws.getRow(3).commit();

    ws.getCell('A4').value = 1.5;
    ws.getCell('A4').numFmt = utils.testValues.numFmt1;
    ws.getCell('A4').border = utils.styles.borders.thin;
    ws.getCell('C4').value = 1.5;
    ws.getCell('C4').numFmt = utils.testValues.numFmt2;
    ws.getCell('C4').border = utils.styles.borders.doubleRed;
    ws.getCell('E4').value = 1.5;
    ws.getCell('E4').border = utils.styles.borders.thickRainbow;
    ws.getRow(4).commit();

    // test fonts and formats
    ws.getCell('A5').value = utils.testValues.str;
    ws.getCell('A5').font = utils.styles.fonts.arialBlackUI14;
    ws.getCell('B5').value = utils.testValues.str;
    ws.getCell('B5').font = utils.styles.fonts.broadwayRedOutline20;
    ws.getCell('C5').value = utils.testValues.str;
    ws.getCell('C5').font = utils.styles.fonts.comicSansUdB16;

    ws.getCell('D5').value = 1.6;
    ws.getCell('D5').numFmt = utils.testValues.numFmt1;
    ws.getCell('D5').font = utils.styles.fonts.arialBlackUI14;

    ws.getCell('E5').value = 1.6;
    ws.getCell('E5').numFmt = utils.testValues.numFmt2;
    ws.getCell('E5').font = utils.styles.fonts.broadwayRedOutline20;

    ws.getCell('F5').value = utils.testValues.date;
    ws.getCell('F5').numFmt = utils.testValues.numFmtDate;
    ws.getCell('F5').font = utils.styles.fonts.comicSansUdB16;
    ws.getRow(5).commit();

    ws.getRow(6).height = 42;
    _.each(utils.styles.alignments, function(alignment, index) {
      var rowNumber = 6;
      var colNumber = index + 1;
      var cell = ws.getCell(rowNumber, colNumber);
      cell.value = alignment.text;
      cell.alignment = alignment.alignment;
    });
    ws.getRow(6).commit();

    if (checkBadAlignments) {
      _.each(utils.styles.badAlignments, function(alignment, index) {
        var rowNumber = 7;
        var colNumber = index + 1;
        var cell = ws.getCell(rowNumber, colNumber);
        cell.value = alignment.text;
        cell.alignment = alignment.alignment;
      });
    }
    ws.getRow(7).commit();

    var row8 = ws.getRow(8);
    row8.height = 40;
    row8.getCell(1).value = 'Blue White Horizontal Gradient';
    row8.getCell(1).fill = utils.styles.fills.blueWhiteHGrad;
    row8.getCell(2).value = 'Red Dark Vertical';
    row8.getCell(2).fill = utils.styles.fills.redDarkVertical;
    row8.getCell(3).value = 'Red Green Dark Trellis';
    row8.getCell(3).fill = utils.styles.fills.redGreenDarkTrellis;
    row8.getCell(4).value = 'RGB Path Gradient';
    row8.getCell(4).fill = utils.styles.fills.rgbPathGrad;
    row8.commit();

    return wb;
  },

  checkTestBook: function(wb, docType, useStyles) {
    var sheetName;
    var checkFormulas, checkMerges, checkStyles, checkBadAlignments;
    var dateAccuracy;
    switch(docType) {
      case 'xlsx':
        sheetName = 'blort';
        checkFormulas = true;
        checkMerges = true;
        checkStyles = useStyles;
        checkBadAlignments = useStyles;
        dateAccuracy = 3;
        break;
      case 'model':
        sheetName = 'blort';
        checkFormulas = true;
        checkMerges = true;
        checkStyles = true;
        checkBadAlignments = false;
        dateAccuracy = 3;
        break;
      case 'csv':
        sheetName = 'sheet1';
        checkFormulas = false;
        checkMerges = false;
        checkStyles = false;
        checkBadAlignments = false;
        dateAccuracy = 1000;
        break;
    }

    expect(wb).to.not.be.undefined;

    var ws = wb.getWorksheet(sheetName);
    expect(ws).to.not.be.undefined;

    expect(ws.getCell('A1').value).to.equal(7);
    expect(ws.getCell('A1').type).to.equal(Excel.ValueType.Number);
    expect(ws.getCell('B1').value).to.equal(utils.testValues.str);
    expect(ws.getCell('B1').type).to.equal(Excel.ValueType.String);
    expect(Math.abs(ws.getCell('C1').value.getTime() - utils.testValues.date.getTime())).to.be.below(dateAccuracy);
    expect(ws.getCell('C1').type).to.equal(Excel.ValueType.Date);

    if (checkFormulas) {
      expect(ws.getCell('D1').value).to.deep.equal(utils.testValues.formulas[0]);
      expect(ws.getCell('D1').type).to.equal(Excel.ValueType.Formula);
      expect(ws.getCell('E1').value).to.deep.equal(utils.testValues.formulas[1]);
      expect(ws.getCell('E1').type).to.equal(Excel.ValueType.Formula);
      expect(ws.getCell('F1').value).to.deep.equal(utils.testValues.hyperlink);
      expect(ws.getCell('F1').type).to.equal(Excel.ValueType.Hyperlink);
      expect(ws.getCell('G1').value).to.equal(utils.testValues.str2);
    } else {
      expect(ws.getCell('D1').value).to.equal(utils.testValues.formulas[0].result);
      expect(ws.getCell('D1').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('E1').value).to.be.null;
      expect(ws.getCell('E1').type).to.equal(Excel.ValueType.Null);
      expect(ws.getCell('F1').value).to.deep.equal(utils.testValues.hyperlink.hyperlink);
      expect(ws.getCell('F1').type).to.equal(Excel.ValueType.String);
      expect(ws.getCell('G1').value).to.equal(utils.testValues.str2);
    }

    // A2:B3
    expect(ws.getCell('A2').value).to.equal(5);
    expect(ws.getCell('A2').type).to.equal(Excel.ValueType.Number);
    expect(ws.getCell('A2').master).to.equal(ws.getCell('A2'));

    if (checkMerges) {
      expect(ws.getCell('A3').value).to.equal(5);
      expect(ws.getCell('A3').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('A3').master).to.equal(ws.getCell('A2'));

      expect(ws.getCell('B2').value).to.equal(5);
      expect(ws.getCell('B2').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('B2').master).to.equal(ws.getCell('A2'));

      expect(ws.getCell('B3').value).to.equal(5);
      expect(ws.getCell('B3').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('B3').master).to.equal(ws.getCell('A2'));

      // C2:D3
      expect(ws.getCell('C2').value).to.be.null;
      expect(ws.getCell('C2').type).to.equal(Excel.ValueType.Null);
      expect(ws.getCell('C2').master).to.equal(ws.getCell('C2'));

      expect(ws.getCell('D2').value).to.be.null;
      expect(ws.getCell('D2').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('D2').master).to.equal(ws.getCell('C2'));

      expect(ws.getCell('C3').value).to.be.null;
      expect(ws.getCell('C3').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('C3').master).to.equal(ws.getCell('C2'));

      expect(ws.getCell('D3').value).to.be.null;
      expect(ws.getCell('D3').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('D3').master).to.equal(ws.getCell('C2'));
    }

    if (checkStyles) {
      expect(ws.getCell('A4').numFmt).to.equal(utils.testValues.numFmt1);
      expect(ws.getCell('A4').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('A4').border).to.deep.equal(utils.styles.borders.thin);
      expect(ws.getCell('C4').numFmt).to.equal(utils.testValues.numFmt2);
      expect(ws.getCell('C4').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('C4').border).to.deep.equal(utils.styles.borders.doubleRed);
      expect(ws.getCell('E4').border).to.deep.equal(utils.styles.borders.thickRainbow);

      // test fonts and formats
      expect(ws.getCell('A5').value).to.equal(utils.testValues.str);
      expect(ws.getCell('A5').type).to.equal(Excel.ValueType.String);
      expect(ws.getCell('A5').font).to.deep.equal(utils.styles.fonts.arialBlackUI14);
      expect(ws.getCell('B5').value).to.equal(utils.testValues.str);
      expect(ws.getCell('B5').type).to.equal(Excel.ValueType.String);
      expect(ws.getCell('B5').font).to.deep.equal(utils.styles.fonts.broadwayRedOutline20);
      expect(ws.getCell('C5').value).to.equal(utils.testValues.str);
      expect(ws.getCell('C5').type).to.equal(Excel.ValueType.String);
      expect(ws.getCell('C5').font).to.deep.equal(utils.styles.fonts.comicSansUdB16);

      expect(Math.abs(ws.getCell('D5').value - 1.6)).to.be.below(0.00000001);
      expect(ws.getCell('D5').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('D5').numFmt).to.equal(utils.testValues.numFmt1);
      expect(ws.getCell('D5').font).to.deep.equal(utils.styles.fonts.arialBlackUI14);

      expect(Math.abs(ws.getCell('E5').value - 1.6)).to.be.below(0.00000001);
      expect(ws.getCell('E5').type).to.equal(Excel.ValueType.Number);
      expect(ws.getCell('E5').numFmt).to.equal(utils.testValues.numFmt2);
      expect(ws.getCell('E5').font).to.deep.equal(utils.styles.fonts.broadwayRedOutline20);

      expect(Math.abs(ws.getCell('F5').value.getTime() - utils.testValues.date.getTime())).to.be.below(dateAccuracy);
      expect(ws.getCell('F5').type).to.equal(Excel.ValueType.Date);
      expect(ws.getCell('F5').numFmt).to.equal(utils.testValues.numFmtDate);
      expect(ws.getCell('F5').font).to.deep.equal(utils.styles.fonts.comicSansUdB16);

      expect(ws.getRow(5).height).to.be.undefined;
      expect(ws.getRow(6).height).to.equal(42);
      _.each(utils.styles.alignments, function(alignment, index) {
        var rowNumber = 6;
        var colNumber = index + 1;
        var cell = ws.getCell(rowNumber, colNumber);
        expect(cell.value).to.equal(alignment.text);
        expect(cell.alignment).to.deep.equal(alignment.alignment);
      });

      if (checkBadAlignments) {
        _.each(utils.styles.badAlignments, function(alignment, index) {
          var rowNumber = 7;
          var colNumber = index + 1;
          var cell = ws.getCell(rowNumber, colNumber);
          expect(cell.value).to.equal(alignment.text);
          expect(cell.alignment).to.be.undefined;
        });
      }

      var row8 = ws.getRow(8);
      expect(row8.height).to.equal(40);
      expect(row8.getCell(1).fill).to.deep.equal(utils.styles.fills.blueWhiteHGrad);
      expect(row8.getCell(2).fill).to.deep.equal(utils.styles.fills.redDarkVertical);
      expect(row8.getCell(3).fill).to.deep.equal(utils.styles.fills.redGreenDarkTrellis);
      expect(row8.getCell(4).fill).to.deep.equal(utils.styles.fills.rgbPathGrad);
    }
  },

  checkTestBookReader: function(filename) {
    var wb = new Excel.stream.xlsx.WorkbookReader();

    // expectations
    var dateAccuracy = 3;

    var deferred = Bluebird.defer();

    wb.on('worksheet', function(ws) {
      console.log('Worksheet', ws.name);
      // Sheet name stored in workbook. Not guaranteed here
      // expect(ws.name).to.equal('blort');
      ws.on('row', function(row) {
        console.log('WB Reader row', row.number)
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
const {
  serializeStyle,
  deserializeStyle,
} = require('../../../lib/utils/style-fast-serialize'); // Adjust the import as needed
// const deserializeStyle = require('../../utils/style-fast-deserialize');

describe('Style Serialization and Deserialization', () => {
  it('should correctly serialize and deserialize font', () => {
    const originalStyle = {
      font: {
        name: 'Arial',
        size: 12,
        family: 2,
        scheme: 'minor',
        charset: 1,
        color: {argb: 'FF0000FF'},
        bold: true,
        italic: false,
        underline: 'single',
        vertAlign: 'superscript',
        strike: true,
        outline: false,
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize border', () => {
    const originalStyle = {
      border: {
        top: {style: 'thin', color: {argb: 'FF000000'}},
        left: {style: 'medium', color: {theme: 1}},
        bottom: {style: 'thick', color: {argb: 'FFFF0000'}},
        right: {style: 'dashed', color: {theme: 2}},
        diagonal: {
          style: 'dotted',
          color: {argb: 'FF00FF00'},
          up: true,
          down: false,
        },
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize alignment', () => {
    const originalStyle = {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
        shrinkToFit: false,
        indent: 1,
        readingOrder: 'rtl',
        textRotation: 45,
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize fill (pattern)', () => {
    const originalStyle = {
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FFFF0000'},
        bgColor: {theme: 3},
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize fill (gradient angle)', () => {
    const originalStyle = {
      fill: {
        type: 'gradient',
        gradient: 'angle',
        degree: 90,
        stops: [
          {position: 0, color: {argb: 'FFFF0000'}},
          {position: 1, color: {theme: 4}},
        ],
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize fill (gradient path)', () => {
    const originalStyle = {
      fill: {
        type: 'gradient',
        gradient: 'path',
        center: {left: 0.5, top: 0.5},
        stops: [
          {position: 0, color: {argb: 'FFFF0000'}},
          {position: 1, color: {theme: 4}},
        ],
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize protection', () => {
    const originalStyle = {
      protection: {
        locked: true,
        hidden: false,
      },
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize number format', () => {
    const originalStyle = {
      numFmt: '#,##0.00',
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });

  it('should correctly serialize and deserialize a full style object', () => {
    const originalStyle = {
      font: {
        name: 'Calibri',
        size: 11,
        color: {theme: 1},
        bold: true,
      },
      border: {
        top: {style: 'thin', color: {argb: 'FF000000'}},
        left: {style: 'thin', color: {argb: 'FF000000'}},
        bottom: {style: 'thin', color: {argb: 'FF000000'}},
        right: {style: 'thin', color: {argb: 'FF000000'}},
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FFFF0000'},
      },
      protection: {
        locked: true,
        hidden: false,
      },
      numFmt: '0.00%',
    };

    const serialized = serializeStyle(originalStyle);
    const deserialized = deserializeStyle(serialized);

    expect(deserialized).to.deep.equal(originalStyle);
  });
});

// style.js

const L1_SEPARATOR = '|';
const L2_SEPARATOR = '>';
const L3_SEPARATOR = '<';
const L4_SEPARATOR = ':';
const L5_SEPARATOR = ';';
const L6_SEPARATOR = '!';
const L7_SEPARATOR = '~';
const L8_SEPARATOR = '^';

function encodeBoolean(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (value === true) {
    return '1';
  }
  if (value === false) {
    return '0';
  }
  return value ? 'true' : 'false';
}

function addSerializedBooleanProperty(obj, property, value) {
  if (value === '1') obj[property] = true;
  if (value === '0') obj[property] = false;
}

function serializeColor(color) {
  if (!color) return '';
  return `${color.argb || ''}${L8_SEPARATOR}${color.theme !== undefined ? color.theme : ''}`;
}

function serializeBorder(border) {
  if (!border) return '';
  return `${border.style || ''}${L4_SEPARATOR}${serializeColor(border.color)}`;
}

function serializeFill(fill) {
  if (!fill) return '';
  if (fill.type === 'pattern') {
    return (
      `p${L3_SEPARATOR}${fill.pattern}${L4_SEPARATOR}${serializeColor(fill.fgColor)}` +
      `${L4_SEPARATOR}${serializeColor(fill.bgColor)}`
    );
  }
  if (fill.type === 'gradient') {
    const stops = fill.stops
      .map(stop => `${stop.position}${L7_SEPARATOR}${serializeColor(stop.color)}`)
      .join(L6_SEPARATOR);

    if (fill.gradient === 'angle') {
      return `ga${L3_SEPARATOR}${fill.degree}${L4_SEPARATOR}${stops}`;
    }
    return `gp${L3_SEPARATOR}${fill.center.left}${L5_SEPARATOR}${fill.center.top}${L4_SEPARATOR}${stops}`;
  }
  return '';
}

function serializeFont(font) {
  return [
    font.name || '',
    font.size || '',
    font.family || '',
    font.scheme || '',
    font.charset || '',
    serializeColor(font.color),
    encodeBoolean(font.bold),
    encodeBoolean(font.italic),
    font.underline || '',
    font.vertAlign || '',
    encodeBoolean(font.strike),
    encodeBoolean(font.outline),
  ].join(L3_SEPARATOR);
}

function serializeBorders(border) {
  const up = border.diagonal && border.diagonal.up;
  const down = border.diagonal && border.diagonal.down;

  return [
    serializeBorder(border.top),
    serializeBorder(border.left),
    serializeBorder(border.bottom),
    serializeBorder(border.right),
    serializeBorder(border.diagonal),
    encodeBoolean(up),
    encodeBoolean(down),
  ].join(L3_SEPARATOR);
}

function serializeAlignment(alignment) {
  return [
    alignment.horizontal || '',
    alignment.vertical || '',
    encodeBoolean(alignment.wrapText),
    encodeBoolean(alignment.shrinkToFit),
    alignment.indent || '',
    alignment.readingOrder || '',
    alignment.textRotation !== undefined ? alignment.textRotation : '',
  ].join(L3_SEPARATOR);
}

function serializeProtection(protection) {
  return [encodeBoolean(protection.locked), encodeBoolean(protection.hidden)].join(L3_SEPARATOR);
}

function serializeStyle(style) {
  const parts = [];

  if (style.font) parts.push(`f${L2_SEPARATOR}${serializeFont(style.font)}`);
  if (style.border) parts.push(`b${L2_SEPARATOR}${serializeBorders(style.border)}`);
  if (style.alignment) parts.push(`a${L2_SEPARATOR}${serializeAlignment(style.alignment)}`);
  if (style.fill) parts.push(`fi${L2_SEPARATOR}${serializeFill(style.fill)}`);
  if (style.protection) parts.push(`p${L2_SEPARATOR}${serializeProtection(style.protection)}`);
  if (style.numFmt) parts.push(`n${L2_SEPARATOR}${style.numFmt}`);

  return parts.join(L1_SEPARATOR);
}

function deserializeColor(str) {
  if (!str) return undefined;
  const [argb, theme] = str.split(L8_SEPARATOR);
  const color = {};
  if (argb) color.argb = argb;
  if (theme !== '') color.theme = Number(theme);
  return Object.keys(color).length ? color : undefined;
}

function deserializeBorder(str) {
  if (!str) return undefined;
  const [style, colorStr] = str.split(L4_SEPARATOR);
  const border = {};
  if (style) border.style = style;
  const color = deserializeColor(colorStr);
  if (color) border.color = color;
  return Object.keys(border).length ? border : undefined;
}

function deserializeFill(str) {
  if (!str) return undefined;

  const [type, rest] = str.split(L3_SEPARATOR);
  if (type === 'p') {
    const [pattern, fgColorStr, bgColorStr] = rest.split(L4_SEPARATOR);
    const fill = {
      type: 'pattern',
      pattern,
    };
    const fgColor = deserializeColor(fgColorStr);
    const bgColor = deserializeColor(bgColorStr);
    if (fgColor) fill.fgColor = fgColor;
    if (bgColor) fill.bgColor = bgColor;
    return fill;
  }
  if (type === 'ga' || type === 'gp') {
    const [degreeOrCenter, stopsStr] = rest.split(L4_SEPARATOR);
    const stops = stopsStr.split(L6_SEPARATOR).map(stop => {
      const [position, colorStr] = stop.split(L7_SEPARATOR);
      return {position: Number(position), color: deserializeColor(colorStr)};
    });

    if (type === 'ga') {
      return {
        type: 'gradient',
        gradient: 'angle',
        degree: Number(degreeOrCenter),
        stops,
      };
    }
    const [left, top] = degreeOrCenter.split(L5_SEPARATOR).map(Number);
    return {
      type: 'gradient',
      gradient: 'path',
      center: {left, top},
      stops,
    };
  }
  return undefined;
}

function deserializeFont(values) {
  const [name, size, family, scheme, charset, color, bold, italic, underline, vertAlign, strike, outline] = values;
  const font = {};

  if (name) font.name = name;
  if (size) font.size = Number(size);
  if (family) font.family = Number(family);
  if (scheme) font.scheme = scheme;
  if (charset) font.charset = Number(charset);

  const deserializedColor = deserializeColor(color);
  if (deserializedColor) font.color = deserializedColor;

  addSerializedBooleanProperty(font, 'bold', bold);
  addSerializedBooleanProperty(font, 'italic', italic);

  if (underline === 'true') {
    font.underline = true;
  } else if (underline === 'false') {
    font.underline = false;
  } else if (underline) {
    font.underline = underline;
  }

  if (vertAlign) font.vertAlign = vertAlign;
  addSerializedBooleanProperty(font, 'strike', strike);
  addSerializedBooleanProperty(font, 'outline', outline);

  return font;
}

function deserializeBorders(values) {
  const [top, left, bottom, right, diagonal, up, down] = values;
  const borders = {};

  if (top !== '') {
    const topBorder = deserializeBorder(top);
    if (topBorder) borders.top = topBorder;
  }

  if (left !== '') {
    const leftBorder = deserializeBorder(left);
    if (leftBorder) borders.left = leftBorder;
  }

  if (bottom !== '') {
    const bottomBorder = deserializeBorder(bottom);
    if (bottomBorder) borders.bottom = bottomBorder;
  }

  if (right !== '') {
    const rightBorder = deserializeBorder(right);
    if (rightBorder) borders.right = rightBorder;
  }

  if (diagonal !== '') {
    const diagonalBorder = deserializeBorder(diagonal);
    if (diagonalBorder) {
      borders.diagonal = diagonalBorder;
      addSerializedBooleanProperty(borders.diagonal, 'up', up);
      addSerializedBooleanProperty(borders.diagonal, 'down', down);
    }
  } else if (up || down) {
    borders.diagonal = {};
    addSerializedBooleanProperty(borders.diagonal, 'up', up);
    addSerializedBooleanProperty(borders.diagonal, 'down', down);
  }

  return borders;
}

function deserializeAlignment(values) {
  const [horizontal, vertical, wrapText, shrinkToFit, indent, readingOrder, textRotation] = values;
  const alignment = {};

  if (horizontal) alignment.horizontal = horizontal;
  if (vertical) alignment.vertical = vertical;
  addSerializedBooleanProperty(alignment, 'wrapText', wrapText);
  addSerializedBooleanProperty(alignment, 'shrinkToFit', shrinkToFit);
  if (indent) alignment.indent = Number(indent);
  if (readingOrder) alignment.readingOrder = readingOrder;
  if (textRotation !== '') {
    alignment.textRotation = textRotation === 'vertical' ? 'vertical' : Number(textRotation);
  }

  return alignment;
}

function deserializeProtection(values) {
  const [locked, hidden] = values;
  const protection = {};
  addSerializedBooleanProperty(protection, 'locked', locked);
  addSerializedBooleanProperty(protection, 'hidden', hidden);
  return protection;
}

function deserializeStyle(serialized) {
  const style = {};
  const parts = serialized.split(L1_SEPARATOR);

  for (const part of parts) {
    const [type, values] = part.split(L2_SEPARATOR);
    const valuesArray = values.split(L3_SEPARATOR);

    switch (type) {
      case 'f':
        style.font = deserializeFont(valuesArray);
        break;
      case 'b':
        style.border = deserializeBorders(valuesArray);
        break;
      case 'a':
        style.alignment = deserializeAlignment(valuesArray);
        break;
      case 'fi':
        style.fill = deserializeFill(values);
        break;
      case 'p':
        style.protection = deserializeProtection(valuesArray);
        break;
      case 'n':
        style.numFmt = values;
        break;
    }
  }

  return style;
}

module.exports = {serializeStyle, deserializeStyle};

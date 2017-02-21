'use strict';

var PromishLib = require('../utils/promish');

function setValue(key, value) {
  switch(key.toLowerCase()) {
    case 'promise':
      PromishLib.Promish = value;
      break;
  }
}

module.exports = setValue;
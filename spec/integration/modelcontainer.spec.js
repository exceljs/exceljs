'use strict';

var expect = require('chai').expect;
var Excel = require('../../excel');
var model = require('./data/testmodel');

describe.skip('ModelContainer', function() {
  it('serializes and deserializes to file properly', function() {
    this.timeout(5000);
    
    // clone model
    var mcModel = JSON.parse(JSON.stringify(model));

    var mc = new Excel.ModelContainer(model);
    return mc.xlsx.writeFile('./spec/out/mc.test.xlsx')
      .then(function() {
        var mc2 = new Excel.ModelContainer();
        return mc2.xlsx.readFile('./spec/out/mc.test.xlsx');
      })
      .then(function(mc2) {
        expect(mc2.model).to.deep.equal(mcModel);
      });
  });
});


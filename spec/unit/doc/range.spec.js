'use strict';

var expect = require('chai').expect;
var Range = require('../../../lib/doc/range');

describe('Dimensions', function() {

  function check(d, range, $range, tl, $t$l, br, $b$r, top, left, bottom, right) {
    expect(d.range).to.equal(range);
    expect(d.$range).to.equal($range);
    expect(d.tl).to.equal(tl);
    expect(d.$t$l).to.equal($t$l);
    expect(d.br).to.equal(br);
    expect(d.$b$r).to.equal($b$r);
    expect(d.top).to.equal(top);
    expect(d.left).to.equal(left);
    expect(d.bottom).to.equal(bottom);
    expect(d.right).to.equal(right);
    expect(d.toString()).to.equal(range);
  }

  it('has a valid default value', function() {
    var d = new Range();
    check(d, 'A1:A1', '$A$1:$A$1', 'A1', '$A$1', 'A1', '$A$1', 1, 1, 1, 1);
  });

  it('constructs as expected', function() {
    // check range + rotations
    check(new Range('B5:D10'), 'B5:D10', '$B$5:$D$10', 'B5', '$B$5', 'D10', '$D$10', 5, 2, 10, 4);
    check(new Range('B10:D5'), 'B5:D10', '$B$5:$D$10', 'B5', '$B$5', 'D10', '$D$10', 5, 2, 10, 4);
    check(new Range('D5:B10'), 'B5:D10', '$B$5:$D$10', 'B5', '$B$5', 'D10', '$D$10', 5, 2, 10, 4);
    check(new Range('D10:B5'), 'B5:D10', '$B$5:$D$10', 'B5', '$B$5', 'D10', '$D$10', 5, 2, 10, 4);

    check(new Range('G7','C16'), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range('C7','G16'), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range('C16','G7'), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range('G16','C7'), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);

    check(new Range(7, 3, 16, 7), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range(16, 3, 7, 7), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range(7, 7, 16, 3), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range(16, 7, 7, 3), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);

    check(new Range([7, 3, 16, 7]), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range([16, 3, 7, 7]), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range([7, 7, 16, 3]), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);
    check(new Range([16, 7, 7, 3]), 'C7:G16', '$C$7:$G$16', 'C7', '$C$7', 'G16', '$G$16', 7, 3, 16, 7);

    check(new Range('B5'), 'B5:B5', '$B$5:$B$5', 'B5', '$B$5', 'B5', '$B$5', 5, 2, 5, 2);
  });


  it('expands properly', function() {
    var d = new Range();

    d.expand(1,1,1,3);
    expect(d.tl).to.equal('A1');
    expect(d.br).to.equal('C1');
    expect(d.toString()).to.equal('A1:C1');

    d.expand(1,3,3,3);
    expect(d.tl).to.equal('A1');
    expect(d.br).to.equal('C3');
    expect(d.toString()).to.equal('A1:C3');
  });

  it('doesn\'t always include the default row/col', function() {
    var d = new Range();

    d.expand(2,2,4,4);
    expect(d.tl).to.equal('B2');
    expect(d.br).to.equal('D4');
    expect(d.toString()).to.equal('B2:D4');
  });

  it('detects intersections', function() {
    var C3F6 = new Range('C3:F6');

    // touching at corners
    expect(C3F6.intersects(new Range('A1:B2'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G1:H2'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A7:B8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G7:H8'))).to.not.be.ok;

    // Adjacent to edges
    expect(C3F6.intersects(new Range('A1:H2'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A1:B8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G1:H8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A7:H8'))).to.not.be.ok;

    // 1 cell margin
    expect(C3F6.intersects(new Range('A1:H1'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A1:A8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G1:G8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A8:G8'))).to.not.be.ok;

    // Adjacent at corners
    expect(C3F6.intersects(new Range('A1:B3'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A1:C2'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('F1:H2'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G1:H3'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A6:B8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('A7:C8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('F7:H8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G6:H8'))).to.not.be.ok;

    // Adjacent at edges
    expect(C3F6.intersects(new Range('A4:B5'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('D1:E2'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('D7:E8'))).to.not.be.ok;
    expect(C3F6.intersects(new Range('G4:H8'))).to.not.be.ok;

    // intersecting at corners
    expect(C3F6.intersects(new Range('A1:C3'))).to.be.ok;
    expect(C3F6.intersects(new Range('F1:H3'))).to.be.ok;
    expect(C3F6.intersects(new Range('A6:C8'))).to.be.ok;
    expect(C3F6.intersects(new Range('F6:H8'))).to.be.ok;

    // slice through middle
    expect(C3F6.intersects(new Range('A4:H5'))).to.be.ok;
    expect(C3F6.intersects(new Range('D1:E8'))).to.be.ok;

    // inside
    expect(C3F6.intersects(new Range('D4:E5'))).to.be.ok;

    // outside
    expect(C3F6.intersects(new Range('A1:H8'))).to.be.ok;
  });
});
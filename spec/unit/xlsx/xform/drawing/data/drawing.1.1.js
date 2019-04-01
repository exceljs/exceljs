var Anchor = require('../../../../../../lib/doc/anchor');

module.exports = {
  "anchors":[
    {
      "range": "A1:C7",
      "tl": new Anchor({"row": 0, "col": 0}),
      "br": new Anchor({"row": 7, "col": 3}),
      "picture": {
        "index": 1,
        "rId": "rId1"
      }
    },
    {
      "range": {
        "tl": new Anchor({"row": 2.5, "col": 5.5}),
        "br": new Anchor({"row": 10.5, "col": 8.5}),
      },
      "tl": new Anchor({"row": 2.5, "col": 5.5}),
      "br": new Anchor({"row": 10.5, "col": 8.5}),
      "picture": {
        "index": 2,
        "rId": "rId2"
      }
    }
  ]
};
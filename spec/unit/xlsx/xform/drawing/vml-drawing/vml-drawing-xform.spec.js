const testXformHelper = require('../../test-xform-helper');

const VmlDrawingXform = verquire('xlsx/xform/drawing/vlm-drawing/vml-drawing-xform');

const options = {
  rels: {
    rId1: {Target: '../media/image1.jpg'},
    rId2: {Target: '../media/image2.jpg'},
  },
  mediaIndex: {image1: 0, 'image1.jpg': 0, image2: 1, 'image2.jpg': 1},
  media: [{}, {}],
};

const expectations = [
  {
    title: 'hfimage',
    create() {
      return new VmlDrawingXform();
    },
    preparedModel: {
      comments: [],
      hfImages: [
        {
          anchor: null,
          editAs: null,
          fill: {
            focussize: '0,0',
            on: 'f',
          },
          id: 'RH',
          imagedata: {
            index: 0,
            rId: 'rId1',
            title: 'pig (1)',
          },
          path: {},
          protection: null,
          shadow: null,
          stroke: {
            on: 'f',
          },
          style: {
            height: '142pt',
            left: '0pt',
            'margin-left': '0pt',
            'margin-top': '0pt',
            position: 'absolute',
            top: '0pt',
            width: '142pt',
          },
          type: 'hfimage',
        },
      ],
    },
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <xml xmlns:oa="urn:schemas-microsoft-com:office:activation"
          xmlns:p="urn:schemas-microsoft-com:office:powerpoint"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml">
          <o:shapelayout v:ext="edit">
              <o:idmap v:ext="edit" data="1" />
          </o:shapelayout>
          <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
              path="m,l,21600r21600,l21600,xe">
              <v:stroke joinstyle="miter" />
              <v:path gradientshapeok="t" o:connecttype="rect" />
          </v:shapetype>
          <v:shape id="RH" o:spid="_x0000_s1025" o:spt="75" alt="pig (1)" type="#_x0000_t75"
              style="position:absolute;left:0pt;top:0pt;margin-left:0pt;margin-top:0pt;height:142pt;width:142pt;"
              filled="f" o:preferrelative="t" stroked="f" coordsize="21600,21600">
              <v:path />
              <v:fill on="f" focussize="0,0" />
              <v:stroke on="f" />
              <v:imagedata o:relid="rId1" o:title="pig (1)" />
              <o:lock v:ext="edit" rotation="t" aspectratio="t" />
          </v:shape>
      </xml>`,
    parsedModel: {
      comments: [],
      hfImages: [
        {
          anchor: null,
          editAs: null,
          fill: {
            focussize: '0,0',
            on: 'f',
          },
          id: 'RH',
          imagedata: {
            rId: 'rId1',
            title: 'pig (1)',
          },
          path: {},
          protection: null,
          shadow: null,
          stroke: {
            on: 'f',
          },
          style: {
            height: '142pt',
            left: '0pt',
            'margin-left': '0pt',
            'margin-top': '0pt',
            position: 'absolute',
            top: '0pt',
            width: '142pt',
          },
          type: 'hfimage',
        },
      ],
    },
    reconciledModel: {
      comments: [],
      hfImages: [
        {
          anchor: null,
          editAs: null,
          fill: {
            focussize: '0,0',
            on: 'f',
          },
          id: 'RH',
          imagedata: {
            rId: 'rId1',
            title: 'pig (1)',
          },
          path: {},
          protection: null,
          shadow: null,
          stroke: {
            on: 'f',
          },
          style: {
            height: '142pt',
            left: '0pt',
            'margin-left': '0pt',
            'margin-top': '0pt',
            position: 'absolute',
            top: '0pt',
            width: '142pt',
          },
          type: 'hfimage',
        },
      ],
    },
    options,
    tests: ['render', 'renderIn', 'parse', 'reconcile'],
  },
];

describe('VmlDrawingXform', () => {
  testXformHelper(expectations);
});

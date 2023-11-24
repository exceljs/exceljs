module.exports = {
  anchors: [
    {
      anchorType: 'xdr:twoCellAnchor',
      range: {
        tl: {
          nativeRow: 4,
          nativeRowOff: 165100,
          nativeCol: 1,
          nativeColOff: 647700,
        },
        br: {
          nativeRow: 10,
          nativeRowOff: 165100,
          nativeCol: 4,
          nativeColOff: 508000,
        },
        editAs: 'oneCell',
      },
      shape: {
        index: 1,
        props: {
          type: 'rect',
          fill: {
            type: 'solid',
            color: {
              theme: 'accent6',
            },
          },
          textBody: {
            vertAlign: 'b',
            paragraphs: [
              {
                alignment: 'l',
                runs: [
                  {
                    font: {size: 11},
                    text: 'Shape1',
                  },
                ],
              },
            ],
          },
        },
      },
    },
    {
      anchorType: 'xdr:twoCellAnchor',
      range: {
        tl: {
          nativeCol: 6,
          nativeColOff: 101600,
          nativeRow: 12,
          nativeRowOff: 63500,
        },
        br: {
          nativeCol: 7,
          nativeColOff: 190500,
          nativeRow: 16,
          nativeRowOff: 165100,
        },
        editAs: 'oneCell',
      },
      picture: null,
      shape: {
        index: 2,
        props: {
          type: 'ellipse',
          fill: {
            type: 'solid',
            color: {
              rgb: 'C651E9',
            },
          },
        },
      },
    },
  ],
};

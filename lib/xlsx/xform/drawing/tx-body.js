module.exports = {
  tag: 'xdr:txBody',
  c: [
    {
      tag: 'a:bodyPr',
      $: {
        vertOverflow: 'clip',
        horzOverflow: 'clip',
        rtlCol: '0',
        anchor: 't',
      },
    },
    {
      tag: 'a:lstStyle',
    },
    {
      tag: 'a:p',
      c: [
        {
          tag: 'a:pPr',
          $: {
            algn: 'l',
          },
        },
        {
          tag: 'a:endParaRPr',
          $: {
            lang: 'en-US',
            sz: '1100',
          },
        },
      ],
    },
  ],
};

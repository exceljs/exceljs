const testXformHelper = require('../../test-xform-helper');

const DataValidationsExtXform = verquire(
  'xlsx/xform/sheet/dv-ext/data-validations-ext-xform'
);

const expectations = [
  {
    title: 'Single element external list type',
    create: () => new DataValidationsExtXform(),
    preparedModel: {
      E1: {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['ExternalSheet!$A$1:$A$10'],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
    <x14:dataValidations count="1"
        xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">
            <x14:formula1>
                <xm:f>ExternalSheet!$A$1:$A$10</xm:f>
            </x14:formula1>
            <xm:sqref>E1</xm:sqref>
        </x14:dataValidation>
    </x14:dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Single element validated by single element external list type',
    create: () => new DataValidationsExtXform(),
    preparedModel: {
      E1: {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['ExternalSheet!$A$1'],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
    <x14:dataValidations count="1"
        xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">
            <x14:formula1>
                <xm:f>ExternalSheet!$A$1</xm:f>
            </x14:formula1>
            <xm:sqref>E1</xm:sqref>
        </x14:dataValidation>
    </x14:dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Multiple element external list type',
    create: () => new DataValidationsExtXform(),
    preparedModel: {
      D1: {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['ExternalSheet2!$B$1:$B$3'],
      },
      E1: {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['ExternalSheet!$A$1:$A$10'],
      },
      E2: {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['ExternalSheet!$A$1:$A$10'],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
    <x14:dataValidations count="2"
        xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">
            <x14:formula1>
                <xm:f>ExternalSheet2!$B$1:$B$3</xm:f>
            </x14:formula1>
            <xm:sqref>D1</xm:sqref>
        </x14:dataValidation>
        <x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">
            <x14:formula1>
                <xm:f>ExternalSheet!$A$1:$A$10</xm:f>
            </x14:formula1>
            <xm:sqref>E1:E2</xm:sqref>
        </x14:dataValidation>
    </x14:dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('DataValidationsExtXform', () => {
  testXformHelper(expectations);
});

// NOTE on ESLint v10 compatibility:
// eslint-config-airbnb-base@15 (its final release; eslintrc-only, declares eslint
// peer <=8) and eslint-plugin-import@2.32.0 (declares eslint peer <=9) predate flat
// config, so pnpm reports unmet-peer warnings against eslint@10. This is expected
// and cosmetic: FlatCompat is the sanctioned bridge for running legacy shareable
// configs under flat-config ESLint, and every airbnb/import rule has been verified
// to load and run correctly on eslint@10.6.x. The durable fix is to move off the
// abandoned airbnb-base (e.g. eslint-plugin-import -> import-x + a maintained flat
// base); that is a behaviour-changing change tracked as a separate follow-up, kept
// out of this migration which only restores the pre-existing lint behaviour.
const {FlatCompat} = require('@eslint/eslintrc');
const js = require('@eslint/js');

const compat = new FlatCompat({
  baseDirectory: __dirname,
  recommendedConfig: js.configs.recommended,
});

module.exports = [
  {
    ignores: ['build/**', 'dist/**', 'out/**', 'spec/manual/public/**'],
  },

  {
    linterOptions: {
      reportUnusedDisableDirectives: 'off',
    },
  },

  ...compat.config({
    parserOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
    },
    extends: ['airbnb-base', 'prettier', 'plugin:n/recommended'],
    env: {
      es6: true,
      mocha: true,
      node: true,
    },
    rules: {
      'arrow-parens': ['error', 'as-needed'],
      'class-methods-use-this': ['off'],
      'comma-dangle': [
        'error',
        {
          arrays: 'always-multiline',
          objects: 'always-multiline',
          imports: 'always-multiline',
          exports: 'always-multiline',
          functions: 'never',
        },
      ],
      'default-case': ['off'],
      'func-names': ['off', 'never'],
      'global-require': ['off'],
      'max-len': ['error', {code: 120, ignoreComments: true, ignoreStrings: true}],
      'no-console': ['error', {allow: ['warn']}],
      'no-continue': ['off'],
      'no-mixed-operators': ['error', {allowSamePrecedence: true}],
      'no-multi-assign': ['off'],
      'no-param-reassign': ['off'],
      'no-path-concat': ['off'],
      'no-plusplus': ['off'],
      'no-prototype-builtins': ['off'],
      'no-restricted-syntax': ['error', 'ForInStatement', 'LabeledStatement', 'WithStatement'],
      'no-return-assign': ['off'],
      'no-trailing-spaces': ['error', {skipBlankLines: true}],
      'no-underscore-dangle': ['off', {allowAfterThis: true, allowAfterSuper: true}],
      'no-unused-vars': ['error', {vars: 'all', args: 'none', ignoreRestSiblings: true}],
      'no-use-before-define': ['error', {variables: false, classes: false, functions: false}],
      'n/no-unsupported-features/es-syntax': ['error', {version: '>=22.21.1', ignores: []}],
      'n/process-exit-as-throw': ['off'],
      'n/no-process-exit': ['off'],
      'object-curly-spacing': ['error', 'never'],
      'object-property-newline': ['off', {allowMultiplePropertiesPerLine: true}],
      'prefer-destructuring': ['warn', {array: false, object: true}],
      'prefer-object-spread': ['off'],
      'prefer-rest-params': ['off'],
      quotes: ['error', 'single'],
      semi: ['error', 'always'],
      'space-before-function-paren': [
        'error',
        {anonymous: 'never', named: 'never', asyncArrow: 'always'},
      ],
      strict: ['off'],
      'import/extensions': ['off'],
      'import/no-unresolved': ['warn'],
      'n/no-missing-require': ['warn'],
      'prefer-arrow-callback': ['warn'],
      radix: ['warn'],
      'arrow-body-style': ['off'],
      'grouped-accessor-pairs': ['warn'],
      'no-promise-executor-return': ['warn'],
    },
  }),

  {
    files: ['spec/**/*.js'],
    languageOptions: {
      globals: {
        verquire: 'writable',
        describe: 'writable',
        expect: 'writable',
        before: 'writable',
        after: 'writable',
        beforeEach: 'writable',
        afterEach: 'writable',
        it: 'writable',
      },
    },
    rules: {
      'no-new': ['off'],
      'max-len': ['off'],
      'brace-style': ['off'],
      'array-bracket-spacing': ['off'],
      'no-sparse-arrays': ['off'],
      'object-property-newline': ['off'],
      'prefer-object-spread': ['off'],
      'no-underscore-dangle': ['off'],
    },
  },

  {
    files: ['spec/browser/**/*.js'],
    languageOptions: {
      globals: {
        window: 'readonly',
        ExcelJS: 'readonly',
      },
    },
  },

  {
    files: ['test/**/*.js'],
    rules: {
      'no-new': ['off'],
      'max-len': ['off'],
      'no-console': ['off'],
      'no-underscore-dangle': ['off'],
      'spaced-comment': ['off'],
    },
  },

  {
    files: ['eslint.config.js', 'playwright.config.js'],
    rules: {
      'import/no-extraneous-dependencies': ['error', {devDependencies: true}],
    },
  },
];

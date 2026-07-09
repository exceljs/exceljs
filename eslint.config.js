const js = require("@eslint/js");
const n = require("eslint-plugin-n");
const eslintConfigPrettier = require("eslint-config-prettier");
const globals = require("globals");

module.exports = [
  js.configs.recommended,
  n.configs["flat/recommended-script"],
  {
    languageOptions: {
      sourceType: "script",
      ecmaVersion: 2022,
      globals: {
        ...globals.node,
        ...globals.mocha,
      },
    },
    rules: {
      "arrow-parens": ["error", "as-needed"],
      "class-methods-use-this": ["off"],
      "comma-dangle": [
        "error",
        {
          arrays: "always-multiline",
          objects: "always-multiline",
          imports: "always-multiline",
          exports: "always-multiline",
          functions: "never",
        },
      ],
      "default-case": ["off"],
      "func-names": ["off", "never"],
      "global-require": ["off"],
      "max-len": [
        "error",
        { code: 120, ignoreComments: true, ignoreStrings: true },
      ],
      "no-console": ["error", { allow: ["warn"] }],
      "no-continue": ["off"],
      "no-mixed-operators": ["error", { allowSamePrecedence: true }],
      "no-multi-assign": ["off"],
      "no-param-reassign": ["off"],
      "no-path-concat": ["off"],
      "no-plusplus": ["off"],
      "no-prototype-builtins": ["off"],
      "no-restricted-syntax": [
        "error",
        "ForInStatement",
        "LabeledStatement",
        "WithStatement",
      ],
      "no-return-assign": ["off"],
      "no-trailing-spaces": ["error", { skipBlankLines: true }],
      "no-underscore-dangle": [
        "off",
        { allowAfterThis: true, allowAfterSuper: true },
      ],
      "no-unused-vars": [
        "error",
        { vars: "all", args: "none", ignoreRestSiblings: true },
      ],
      "no-use-before-define": [
        "error",
        { variables: false, classes: false, functions: false },
      ],
      "n/no-unpublished-require": ["off"],
      "n/no-process-exit": ["off"],
      "n/no-unsupported-features/es-syntax": [
        "error",
        { version: ">=10.0.0", ignores: [] },
      ],
      "n/no-unsupported-features/es-builtins": [
        "error",
        { version: ">=10.0.0", ignores: [] },
      ],
      "n/no-unsupported-features/node-builtins": [
        "error",
        { version: ">=10.0.0", ignores: [] },
      ],
      "n/no-missing-require": ["error"],
      "object-curly-spacing": ["error", "never"],
      "object-property-newline": [
        "off",
        { allowMultiplePropertiesPerLine: true },
      ],
      "prefer-destructuring": ["warn", { array: false, object: true }],
      "prefer-object-spread": ["off"],
      "prefer-rest-params": ["off"],
      quotes: ["error", "single"],
      semi: ["error", "always"],
      "space-before-function-paren": [
        "error",
        { anonymous: "never", named: "never", asyncArrow: "always" },
      ],
      strict: ["off"],
    },
  },
  eslintConfigPrettier,
  {
    files: ["spec/**"],
    languageOptions: {
      globals: {
        expect: "readonly",
        verquire: "readonly",
      },
    },
  },
  {
    files: ["build.js", "eslint.config.js"],
    rules: {
      "n/no-unsupported-features/es-syntax": ["off"],
      "n/no-unsupported-features/node-builtins": ["off"],
      "n/no-missing-require": ["off"],
      "n/no-unpublished-require": ["off"],
    },
  },
  {
    ignores: [
      "build/**",
      "dist/**",
      "out/**",
      "spec/manual/public/**",
      "node_modules/**",
      ".grunt/**",
    ],
  },
];

module.exports = {
  parser: '@typescript-eslint/parser',
  extends: ['@typescript-eslint/recommended', 'prettier'],
  plugins: ['@typescript-eslint', 'prettier'],
  parserOptions: {
    ecmaVersion: 2017,
    sourceType: 'module',
    project: './tsconfig.json',
  },
  rules: {
    'prettier/prettier': 'error',
    'no-unused-vars': 'off', // Turn off base rule
    '@typescript-eslint/no-unused-vars': [
      'error',
      {
        varsIgnorePattern: '^main$',
        argsIgnorePattern: '^_',
      },
    ],
    // Remove the duplicate rules above

    '@typescript-eslint/explicit-function-return-type': 'warn',
    '@typescript-eslint/prefer-const': 'error',
    '@typescript-eslint/no-explicit-any': 'error',
    'no-console': 'warn',
    'prefer-const': 'error',
    'no-var': 'error',
    '@typescript-eslint/naming-convention': [
      'error',
      {
        selector: 'variableLike',
        format: ['camelCase'],
      },
      {
        selector: 'function',
        format: ['camelCase'],
      },
      {
        selector: 'interface',
        format: ['PascalCase'],
        prefix: ['I'],
      },
    ],
  },
  env: {
    browser: false,
    node: false,
    es6: true,
  },
  globals: {
    ExcelScript: 'readonly',
    OfficeScript: 'readonly',
    console: 'readonly',
  },
};

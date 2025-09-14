module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: ['eslint:recommended', 'plugin:@typescript-eslint/recommended', 'prettier'],
  env: { es2021: true, node: true, browser: true },
  ignorePatterns: ['dist', '**/*.d.ts'],
  rules: {
    '@typescript-eslint/no-explicit-any': 'off'
  }
};


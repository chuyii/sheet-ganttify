import js from '@eslint/js';
import { globalIgnores } from 'eslint/config';
import eslintConfigPrettier from 'eslint-config-prettier/flat';
import importPlugin from 'eslint-plugin-import';
import globals from 'globals';
import neostandard from 'neostandard';
import tseslint from 'typescript-eslint';

export default tseslint.config([
  globalIgnores([
    'template/**/*',
    'template-ui/**/*',
    '**/node_modules/*',
    'build/*',
    '**/dist/*',
    '**/testing',
  ]),
  {
    languageOptions: {
      ...tseslint.configs.base.languageOptions,
      globals: {
        ...globals.browser,
        ...globals.node,
      },
      parserOptions: {
        projectService: {
          allowDefaultProject: ['*.mjs'],
        },
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },
  ...neostandard(),
  js.configs.recommended,
  {
    files: ['**/*.ts'],
    extends: [tseslint.configs.recommendedTypeChecked],
    rules: {
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
    },
  },
  {
    extends: [
      importPlugin.flatConfigs.recommended,
      importPlugin.flatConfigs.typescript,
    ],
    languageOptions: {
      ecmaVersion: 'latest',
    },
    settings: {
      'import/resolver': {
        typescript: true,
      },
    },
    rules: {
      'import/no-named-as-default-member': 'off',
      'import/order': [
        'warn',
        {
          named: true,
          alphabetize: { order: 'asc' },
        },
      ],
    },
  },
  eslintConfigPrettier,
]);

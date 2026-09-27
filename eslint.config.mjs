import tsPlugin from "@typescript-eslint/eslint-plugin";
import tsParser from "@typescript-eslint/parser";
import officeAddins from "eslint-plugin-office-addins";
import globals from "globals";

/**
 * ESLint 9 flat config.
 *
 * Two things the shared office-addins config leaves to the project:
 *  - it declares no globals, so `Excel`, `DOMParser` and `describe` would all
 *    trip no-undef;
 *  - its parser block does not reliably reach this project's .ts/.tsx files, so
 *    the TypeScript parser is wired up explicitly below.
 */
export default [
  {
    ignores: ["dist/**", "public/**", "node_modules/**", "coverage/**"],
  },

  ...officeAddins.configs.react,

  // Everything in the add-in runs in the task pane's browser sandbox.
  {
    files: ["**/*.{ts,tsx,js,mjs,cjs}"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        ...globals.browser,
        // Injected by office.js at runtime rather than imported.
        Excel: "readonly",
        Office: "readonly",
        OfficeExtension: "readonly",
        // Provided by webpack for hot module replacement.
        module: "readonly",
        require: "readonly",
      },
    },
  },

  {
    files: ["**/*.{ts,tsx}"],
    plugins: { "@typescript-eslint": tsPlugin },
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: "latest",
        sourceType: "module",
        ecmaFeatures: { jsx: true },
      },
    },
    rules: {
      // The base rule cannot see TypeScript type positions, so it reports the
      // parameter names in a function *type* as unused variables.
      "no-unused-vars": "off",
      "@typescript-eslint/no-unused-vars": [
        "error",
        { args: "after-used", argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],
    },
  },

  // Build tooling and unit tests run in Node, not in Excel.
  {
    files: ["**/*.js", "**/*.mjs", "**/*.cjs"],
    languageOptions: {
      sourceType: "commonjs",
      globals: { ...globals.node },
    },
  },

  {
    files: ["**/*.mjs"],
    languageOptions: { sourceType: "module" },
  },

  {
    files: ["tests/**/*.ts"],
    languageOptions: {
      globals: { ...globals.jest, ...globals.node },
    },
  },
];

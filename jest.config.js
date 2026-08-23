/**
 * Unit tests cover the pure logic only - grid transforms, header planning,
 * formula translation, mapping. Anything that talks to Office.js is verified by
 * sideloading the add-in into Excel (see README, "Testing").
 */
module.exports = {
  testEnvironment: "node",
  roots: ["<rootDir>/tests"],
  transform: {
    "^.+\\.tsx?$": [
      "ts-jest",
      {
        // The app builds as ES modules; Jest needs CommonJS.
        tsconfig: { module: "commonjs", target: "ES2019", lib: ["ES2020", "DOM"] },
      },
    ],
  },
  collectCoverageFrom: ["src/taskpane/**/*.ts", "!src/taskpane/**/*.d.ts"],
};

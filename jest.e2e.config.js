module.exports = {
  preset: 'jest-puppeteer',
  testEnvironment: 'jest-environment-puppeteer',
  testMatch: ['**/e2e/**/*.test.ts'],
  transform: {
    '^.+\\.ts$': 'ts-jest',
  },
  testTimeout: 120000, // 2 minutes for E2E tests
  setupFilesAfterEnv: ['<rootDir>/e2e/setup.ts'],
  globals: {
    'ts-jest': {
      tsconfig: {
        esModuleInterop: true,
        types: ['jest', 'puppeteer', './e2e/global.d.ts']
      },
    },
  },
};
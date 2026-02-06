module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jest-environment-jsdom',
  roots: ['<rootDir>/tests'],
  transform: {
    '^.+\\.tsx?$': ['ts-jest', {
      tsconfig: {
        // Override tsconfig for tests — ts-jest needs these
        target: 'es2017',
        module: 'commonjs',
        moduleResolution: 'node',
        jsx: 'react',
        esModuleInterop: true,
        allowSyntheticDefaultImports: true,
        strict: false,
        noImplicitAny: false,
        skipLibCheck: true,
        types: ['jest'],
        lib: ['es2017', 'dom'],
      },
    }],
  },
  moduleNameMapper: {
    '\\.(css|scss)$': '<rootDir>/tests/__mocks__/styleMock.js',
    '^spfx-toolkit/(.*)$': '<rootDir>/node_modules/spfx-toolkit/$1',
    '^@store/(.*)$': '<rootDir>/src/libraries/spSearchStore/$1',
    '^@interfaces/(.*)$': '<rootDir>/src/libraries/spSearchStore/interfaces/$1',
    '^@services/(.*)$': '<rootDir>/src/libraries/spSearchStore/services/$1',
    '^@providers/(.*)$': '<rootDir>/src/libraries/spSearchStore/providers/$1',
    '^@registries/(.*)$': '<rootDir>/src/libraries/spSearchStore/registries/$1',
    '^@orchestrator/(.*)$': '<rootDir>/src/libraries/spSearchStore/orchestrator/$1',
    '^@webparts/(.*)$': '<rootDir>/src/webparts/$1',
  },
  testMatch: ['**/*.test.ts', '**/*.test.tsx'],
  collectCoverageFrom: [
    'src/libraries/**/*.ts',
    '!src/libraries/**/index.ts',
    '!src/**/*.d.ts',
  ],
};

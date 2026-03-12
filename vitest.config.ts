import { defineConfig } from 'vitest/config';
import vue from '@vitejs/plugin-vue';

export default defineConfig({
  plugins: [vue()],
  test: {
    projects: [
      {
        test: {
          name: 'node',
          environment: 'node',
          include: ['tests/**/*.test.ts', 'playground/tests/**/*.test.ts'],
          exclude: ['tests/browser/**/*.browser.test.ts', 'playground/tests/**/*.browser.test.ts'],
        },
      },
      {
        test: {
          name: 'browser',
          include: ['tests/browser/**/*.browser.test.ts', 'playground/tests/**/*.browser.test.ts'],
          environment: 'jsdom',
          setupFiles: ['./tests/browser/setup.ts'],
        },
      },
    ],
  },
});

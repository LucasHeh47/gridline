import { readFileSync } from 'node:fs';
import { defineConfig } from 'vite';

const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')) as { version: string };

export default defineConfig({
  base: './',
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    target: 'chrome130',
    chunkSizeWarningLimit: 2000,
  },
  server: { port: 5199, strictPort: true },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'node',
    // Loading ExcelJS for the first file test can take a while when every test file runs at once
    testTimeout: 20000,
  },
} as any);

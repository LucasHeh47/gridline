// Bundles the Electron main and preload scripts (TypeScript -> CommonJS) with Rolldown.
import { rolldown } from 'rolldown';

const entries = [
  { input: 'electron/main.ts', file: 'dist-electron/main.js' },
  { input: 'electron/preload.ts', file: 'dist-electron/preload.js' },
];

for (const e of entries) {
  const bundle = await rolldown({
    input: e.input,
    platform: 'node',
    external: ['electron', /^node:/],
  });
  await bundle.write({ file: e.file, format: 'cjs', sourcemap: false });
  await bundle.close();
}
console.log('Built Electron main + preload');

// Development: starts the Vite dev server (hot reload) and launches Electron pointed at it.
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { createServer } from 'vite';

const require = createRequire(import.meta.url);

const server = await createServer({ configFile: 'vite.config.mts' });
await server.listen();
const url = server.resolvedUrls?.local?.[0] ?? 'http://localhost:5199/';

await import('./build-electron.mjs');

const electronPath = require('electron');
const child = spawn(electronPath, ['.', ...process.argv.slice(2)], {
  stdio: 'inherit',
  env: { ...process.env, VITE_DEV_SERVER_URL: url },
});
child.on('close', async (code) => {
  await server.close();
  process.exit(code ?? 0);
});

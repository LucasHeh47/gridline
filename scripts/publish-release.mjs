// Publishes the built files in release/ as ONE GitHub release, using the GitHub CLI (gh) and GH_TOKEN.
// Uploading from a single step avoids electron-builder's parallel uploads, which can create two
// releases for the same tag.
//
//   node scripts/publish-release.mjs               regular release (marked as the latest)
//   node scripts/publish-release.mjs --prerelease  development build
import { spawnSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync } from 'node:fs';

const version = JSON.parse(readFileSync('package.json', 'utf8')).version;
const tag = `v${version}`;
const prerelease = process.argv.includes('--prerelease');
const repo = process.env.GITHUB_REPOSITORY ? ['--repo', process.env.GITHUB_REPOSITORY] : [];

const files = readdirSync('release')
  .filter((f) => (f.startsWith('Gridline-') && f.includes(version) && /\.(exe|blockmap)$/.test(f)) || (/\.yml$/.test(f) && f !== 'builder-debug.yml'))
  .map((f) => `release/${f}`);
if (!files.some((f) => f.endsWith('.yml')) || !files.some((f) => f.endsWith('.exe'))) {
  console.error(`Nothing to publish for ${version}: build first (files found: ${files.join(', ') || 'none'})`);
  process.exit(1);
}
const notes = existsSync('build/release-notes.md') ? ['--notes-file', 'build/release-notes.md'] : ['--notes', `Gridline ${version}`];

function gh(args) {
  console.log(`> gh ${args.join(' ')}`);
  const r = spawnSync('gh', [...args, ...repo], { stdio: 'inherit', shell: process.platform === 'win32' });
  if (r.status !== 0) process.exit(r.status ?? 1);
}

const exists = spawnSync('gh', ['release', 'view', tag, ...repo], { stdio: 'ignore', shell: process.platform === 'win32' }).status === 0;
if (exists) {
  gh(['release', 'upload', tag, ...files, '--clobber']);
  gh(['release', 'edit', tag, '--title', `Gridline ${version}`, ...notes, ...(prerelease ? ['--prerelease'] : ['--latest'])]);
} else {
  gh(['release', 'create', tag, ...files, '--verify-tag', '--title', `Gridline ${version}`, ...notes, ...(prerelease ? ['--prerelease', '--latest=false'] : ['--latest'])]);
}
console.log(`Published ${tag}: ${files.map((f) => f.slice('release/'.length)).join(', ')}`);

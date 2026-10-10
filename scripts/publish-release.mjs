// Publishes the built files in release/ as ONE GitHub release, using the GitHub CLI (gh) and GH_TOKEN.
// Uploading from a single step avoids electron-builder's parallel uploads, which can create two
// releases for the same tag.
//
//   node scripts/publish-release.mjs               regular release for an existing tag vX.Y.Z (marked latest)
//   node scripts/publish-release.mjs --prerelease  development build; the tag is created with the release
//   add --dry-run to print the gh commands without running them
import { spawnSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync } from 'node:fs';

const version = JSON.parse(readFileSync('package.json', 'utf8')).version;
const tag = `v${version}`;
const prerelease = process.argv.includes('--prerelease');
const dryRun = process.argv.includes('--dry-run');
const repo = process.env.GITHUB_REPOSITORY ? ['--repo', process.env.GITHUB_REPOSITORY] : [];

const files = readdirSync('release')
  .filter((f) => (f.startsWith('Gridline-') && f.includes(version) && /\.(exe|blockmap)$/.test(f)) || (/\.yml$/.test(f) && f !== 'builder-debug.yml'))
  .map((f) => `release/${f}`);
if (!files.some((f) => f.endsWith('.yml')) || !files.some((f) => f.endsWith('.exe'))) {
  console.error(`Nothing to publish for ${version}: build first (files found: ${files.join(', ') || 'none'})`);
  process.exit(1);
}
const notes = existsSync('build/release-notes.md') ? ['--notes-file', 'build/release-notes.md'] : ['--notes', `Gridline ${version}`];

// gh is an .exe, so no shell is needed. (A shell would split arguments that contain spaces.)
function run(args, quiet = false) {
  if (dryRun) {
    console.log(`[dry run] gh ${args.map((a) => (/\s/.test(a) ? `"${a}"` : a)).join(' ')}`);
    return 0;
  }
  if (!quiet) console.log(`> gh ${args.join(' ')}`);
  return spawnSync('gh', [...args, ...repo], { stdio: quiet ? 'ignore' : 'inherit' }).status ?? 1;
}

function gh(args) {
  const status = run(args);
  if (status !== 0) process.exit(status);
}

const exists = !dryRun && run(['release', 'view', tag], true) === 0;
if (exists) {
  gh(['release', 'upload', tag, ...files, '--clobber']);
  gh(['release', 'edit', tag, '--title', `Gridline ${version}`, ...notes, ...(prerelease ? ['--prerelease'] : ['--latest'])]);
} else if (prerelease) {
  // Create the tag together with the release, so a failed build never leaves a tag without a release
  const target = process.env.GITHUB_SHA ? ['--target', process.env.GITHUB_SHA] : [];
  gh(['release', 'create', tag, ...files, ...target, '--title', `Gridline ${version}`, ...notes, '--prerelease', '--latest=false']);
} else {
  gh(['release', 'create', tag, ...files, '--verify-tag', '--title', `Gridline ${version}`, ...notes, '--latest']);
}
console.log(`${dryRun ? 'Would publish' : 'Published'} ${tag}: ${files.map((f) => f.slice('release/'.length)).join(', ')}`);

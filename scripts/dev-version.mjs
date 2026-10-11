// Sets package.json to the next development version, e.g. 1.0.3-beta.<build>.
// While package.json names a version that hasn't been released yet (no v1.0.3 tag), builds are
// previews of that version; once it's released, they become previews of the next one (1.0.4-beta.N).
// "beta" lets people on the Developer branch move on to the next regular release when it ships.
import { execSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';

const build = process.argv[2];
if (!/^\d+$/.test(build ?? '')) {
  console.error('Usage: node scripts/dev-version.mjs <build number>');
  process.exit(1);
}
const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const base = pkg.version.split('-')[0];
const [major, minor, patch] = base.split('.').map(Number);
let released = false;
try {
  released = execSync(`git tag -l v${base}`, { encoding: 'utf8' }).trim() === `v${base}`;
} catch {
  // not a git checkout: assume released
  released = true;
}
pkg.version = released ? `${major}.${minor}.${patch + 1}-beta.${build}` : `${base}-beta.${build}`;
writeFileSync('package.json', JSON.stringify(pkg, null, 2) + '\n');
console.log(pkg.version);

// Sets package.json to the next development version, e.g. 1.0.2 -> 1.0.3-beta.<build>.
// "beta" lets people on the Developer branch move on to the next regular release when it ships.
import { readFileSync, writeFileSync } from 'node:fs';

const build = process.argv[2];
if (!/^\d+$/.test(build ?? '')) {
  console.error('Usage: node scripts/dev-version.mjs <build number>');
  process.exit(1);
}
const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const [major, minor, patch] = pkg.version.split('-')[0].split('.').map(Number);
pkg.version = `${major}.${minor}.${patch + 1}-beta.${build}`;
writeFileSync('package.json', JSON.stringify(pkg, null, 2) + '\n');
console.log(pkg.version);

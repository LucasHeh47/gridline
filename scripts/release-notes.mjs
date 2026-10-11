// Writes build/release-notes.md, which electron-builder uses as the GitHub release text.
// Regular releases use their CHANGELOG.md section; development builds list the recent commits.
import { execSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';

const version = JSON.parse(readFileSync('package.json', 'utf8')).version;
let notes;
if (version.includes('-')) {
  let since = '';
  try {
    since = execSync('git describe --tags --abbrev=0 --match "v[0-9]*.[0-9]*.[0-9]" --exclude "v*-*" HEAD', { encoding: 'utf8' }).trim();
  } catch {
    // no regular release yet
  }
  const log = execSync(`git log ${since ? `${since}..HEAD` : '-15'} --no-merges --format="- %s"`, { encoding: 'utf8' }).trim();
  // The upcoming version's changelog section, when it has been written already
  const upcoming = readFileSync('CHANGELOG.md', 'utf8')
    .split(/^## /m)
    .find((s) => s.startsWith(`${version.split('-')[0]}\n`) || s.startsWith(`${version.split('-')[0]}\r\n`));
  notes = [
    `Development build ${version} from the \`develop\` branch.`,
    '',
    'It has the newest changes and may have bugs. To go back to regular releases, turn off **Settings → Updates → Developer branch**.',
    '',
    upcoming ? `What's in it:\n\n${upcoming.split(/\r?\n/).slice(1).join('\n').trim()}\n` : '',
    log ? `Changes since ${since || 'the last release'}:\n\n${log}` : '',
  ].join('\n');
} else {
  const sections = readFileSync('CHANGELOG.md', 'utf8').split(/^## /m);
  const section = sections.find((s) => s.startsWith(`${version}\n`) || s.startsWith(`${version}\r\n`));
  notes = section ? section.split(/\r?\n/).slice(1).join('\n').trim() : `Gridline ${version}`;
}
writeFileSync('build/release-notes.md', notes.trim() + '\n');
console.log(`Release notes for ${version}:\n${notes}`);

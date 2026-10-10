# Developing Gridline

Gridline is an [Electron](https://www.electronjs.org/) app written in TypeScript. The spreadsheet engine, file formats and UI are all in this repository, with no UI framework.

## Setup

You need [Node.js](https://nodejs.org/) 22 and Git. Then:

```
git clone https://github.com/LucasHeh47/gridline.git
cd gridline
npm install
```

| Command | What it does |
| --- | --- |
| `npm run dev` | Runs the app with hot reload |
| `npm start` | Builds, then runs the production app |
| `npm test` | Runs the test suite (Vitest) |
| `npm run typecheck` | Type-checks the app and the Electron code |
| `npm run build` | Builds `dist/` (UI) and `dist-electron/` (main + preload) |
| `npm run dist` | Builds the Windows installer and portable exe into `release/` |
| `npm run release` | Builds, then publishes one GitHub release with the [GitHub CLI](https://cli.github.com/) (needs `GH_TOKEN`; normally done by CI) |

## Project layout

| Path | Contents |
| --- | --- |
| `src/core/` | The spreadsheet itself, with no UI: workbook model, formula parser and engine, 250+ functions, number formats, xlsx/csv import and export, undo history, Advanced format, presets |
| `src/core/formula/functions/` | Functions by category. `nfl.ts`, `money.ts` and `world.ts` hold the live-data and personal-finance functions |
| `src/core/data/` | The live-data downloader (`external.ts`) and NFL data handling |
| `src/core/config.ts` | Options the core reads (date order, CSV format, cache times…), set from Settings |
| `src/ui/` | The interface: grid canvas, ribbon, Home page, settings, panels, dialogs |
| `src/ui/settings/schema.ts` | Every setting, with the code that applies it |
| `electron/` | Main process (windows, files, downloads, updates) and the preload bridge |
| `tests/` | Unit tests. Live-data tests use canned responses, never the network |
| `docs/` | User documentation |

**Adding a live-data source:**
1. Define the function with `external.get(resource)`; see `world.ts` for a compact example.
2. Add the host to `DATA_HOSTS` in `electron/main.ts`. The desktop app downloads only from allow-listed HTTPS hosts.
3. Add tests with a fake fetcher.

## Branches

| Branch | Purpose |
| --- | --- |
| `main` | Released versions only. Every commit on `main` is a release, tagged `vX.Y.Z`. |
| `develop` | Ongoing work. Every push builds a development release for people on the Developer branch. |

Work goes into `develop` (directly or through pull requests). A release merges `develop` into `main`.

## Versions

Gridline uses `MAJOR.MINOR.PATCH` versions. Each update bumps the version, usually the patch number: 1.0.1 → 1.0.2.

Development builds are numbered automatically: with `main` at `1.0.2`, builds from `develop` are `1.0.3-beta.<build number>`. The `beta` label matters: it lets people on the Developer branch move on to the next regular release when it comes out.

## Releasing a new version

1. On `develop`, set `version` in `package.json` (for example `1.0.3`) and rename the `## Unreleased` section at the top of `CHANGELOG.md` to `## 1.0.3`. Users see that section as "What's new" in the app and in the GitHub release.
2. Commit, then merge into `main` and tag it:
   ```
   git checkout main
   git merge --ff-only develop
   git tag -a v1.0.3 -m "Gridline 1.0.3"
   git push origin main v1.0.3
   git checkout develop
   git push origin develop
   ```
3. The **Release** workflow (`.github/workflows/release.yml`) tests, builds and publishes the installer, the portable exe and `latest.yml` to GitHub Releases. Installed copies find the new version the next time they start.

## Development builds

The **Development build** workflow (`.github/workflows/dev-build.yml`) runs on every push to `develop`. It:

1. runs the tests;
2. sets a `-beta.N` version;
3. tags the commit;
4. publishes a GitHub **pre-release**.

Commits that are already a regular release are skipped.

To receive these builds, turn on **Settings → Updates → Developer branch** in the app.

## Testing updates locally

To see the update prompt without publishing anything:

1. Build an installer with `npm run dist`.
2. Copy it and `release/latest.yml` into a folder. In the copy of `latest.yml`, raise `version` (for example `1.0.9`).
3. Serve that folder, for example with `python -m http.server 8123`.
4. Start the packaged app with `GRIDLINE_UPDATE_FEED=http://127.0.0.1:8123/`. Only localhost feeds are accepted.
   - If Gridline is already running, also pass `--user-data-dir=<some empty folder>`, because only one copy may run per profile.

Don't click **Restart and update** with a test feed unless you want that build installed on your machine.

## Other environment variables (testing)

| Variable | Effect |
| --- | --- |
| `GRIDLINE_TEMPLATE=<id>` | Opens a preset at startup, for example `budget` or `nfl` |
| `GRIDLINE_OPEN=<path>` | Opens a file at startup |
| `GRIDLINE_SCREENSHOT=<file.png>` | Saves a screenshot of the first window and quits (`GRIDLINE_SCREENSHOT_DELAY` sets the wait in ms) |

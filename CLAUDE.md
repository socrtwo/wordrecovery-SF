# CLAUDE.md

Recovers text from damaged, corrupted, or deleted Microsoft Word documents
(`.docx`, `.doc`, `.rtf`, `.odt`, renamed temp files, raw byte blobs). The
canonical engine is **`web/app.js`** — it runs five independent recovery
methods in parallel and ranks the results by quality. Desktop builds wrap
the web frontend with **Tauri 2**; mobile builds wrap it with **Capacitor
6**. Everything runs locally.

## Repo map

- `web/` — the canonical frontend + recovery engine (`web/app.js`).
  Edits go here unless the task is specifically about the desktop /
  mobile wrappers.
- `src-tauri/` — Tauri 2 wrapper for desktop builds (Windows `.msi` /
  `.exe`, macOS `.dmg` Intel + Apple Silicon, Linux `.AppImage` / `.deb`
  / `.rpm`).
- `mobile/` — Capacitor 6 wrapper for Android `.apk` and iOS `.ipa`.
- `socrtwo/` — additional shared modules / assets.
- `Version_3.*-source.zip`, `Version_3.0-alpha-source.7z`,
  `WordRecovery_source_*.zip` — frozen historical source snapshots from
  prior versions. **Read-only — do not modify.**
- `readme.txt` — legacy text README (alongside the modernized
  `README.md`).
- `.github/workflows/` — `pages.yml` (deploy `web/` to Pages on push to
  `main`), `release.yml` (build native installers on `v*` tag).

The five recovery methods (in `web/app.js`):
1. **Standard parse** (JSZip) — extract from `word/document.xml`.
2. **Byte-level ZIP recovery** — `PK\x03\x04` scan + ImmortalInflate with
   brute-force shift sweep (offsets 0..47).
3. **DOCX XML tag fixes** — apply the original SourceForge tool's
   `InvalidTags` → `ValidTags` substitutions.
4. (See `web/app.js` for the remaining two methods.)
5. (As above.)

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. Pages and Release pipelines fire from `main` only.

## Releasing

- Push a `v*` tag to `main` (or use Actions → Release → workflow_dispatch)
  to build:
  - **Desktop** via Tauri: `.msi` + `.exe` (Windows), `.dmg` (macOS
    universal), `.AppImage` + `.deb` + `.rpm` (Linux).
  - **Mobile** via Capacitor: signed `.apk`, unsigned `.ipa` (sideload
    via Xcode / AltStore).
  - **Web**: PWA bundle.

## Verifying changes

- Web frontend / engine: serve `web/` locally and exercise the five
  methods on each supported input format (`.docx`, `.doc`, `.rtf`,
  `.odt`, renamed binary).
- Confirm the **ranking** still picks the highest-quality output — a
  recovery regression often shows up as "method 2 ranks higher than
  method 1 on a healthy file."
- Tauri: `cd src-tauri && cargo tauri dev` for local desktop testing.
- Capacitor: `cd mobile && npx cap run android` for Android emulator
  testing.

## Gotchas

- **`Version_*-source.{zip,7z}` are frozen snapshots.** Don't unzip and
  edit them — every version of the historical source is preserved
  deliberately.
- The five methods must run in parallel and must not throw. A method that
  fails should produce an empty / low-quality result, not an exception
  that breaks the ranking.
- ImmortalInflate is the "never-throws" DEFLATE decoder. Raising an
  exception on bad data is a bug.
- The brute-force shift sweep (offsets 0..47) is intentional and load-
  bearing for the byte-level path. Don't "optimize" it away.
- Mobile iOS `.ipa` is unsigned — signing is the user's responsibility.
  Don't hard-code signing identifiers.

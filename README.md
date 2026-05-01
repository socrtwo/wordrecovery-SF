<!--MODERNIZED:v3.1-->
# Word Recovery

> Recover text from damaged, corrupted, or deleted Microsoft Word documents.
> Modernized, multi-platform rebuild of the original SourceForge project.

[![Live web app](https://img.shields.io/badge/web%20app-live-ff2e93?style=for-the-badge)](https://socrtwo.github.io/wordrecovery-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/wordrecovery-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/wordrecovery-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/wordrecovery-SF?style=for-the-badge&color=22d3ee)](https://github.com/socrtwo/wordrecovery-SF/blob/main/LICENSE)

🌐 **Live web app:** <https://socrtwo.github.io/wordrecovery-SF/>
📦 **Native installers:** [Releases](https://github.com/socrtwo/wordrecovery-SF/releases)
📂 **Source:** [socrtwo/wordrecovery-SF](https://github.com/socrtwo/wordrecovery-SF)

---

## What it does

Open a `.docx`, `.doc`, `.rtf`, `.odt`, a renamed temp file, or any blob of bytes you suspect has Word text inside, and the app runs five recovery methods in parallel and shows the best result. Everything happens locally — nothing is uploaded.

## Platforms

| Platform | Format | Where to get it |
|----------|--------|-----------------|
| **Web (PWA)** | Browser, installable | <https://socrtwo.github.io/wordrecovery-SF/> |
| **Windows** | `.msi` and `.exe` (NSIS) | Releases page |
| **macOS** | `.dmg` (Intel + Apple Silicon) | Releases page |
| **Linux** | `.AppImage`, `.deb`, `.rpm` | Releases page |
| **Android** | `.apk` | Releases page |
| **iOS** | unsigned `.ipa` (sideload via Xcode/AltStore) | Releases page |

The desktop builds wrap the same web frontend with [Tauri 2](https://tauri.app); mobile builds wrap it with [Capacitor 6](https://capacitorjs.com).

## How it works

The recovery engine ([`web/app.js`](web/app.js)) runs **five independent methods** and ranks the results by quality:

1. **Standard parse (JSZip).** Treats the file as a normal Office package and extracts text from `word/document.xml`.
2. **Byte-level ZIP recovery.** Walks the file looking for `PK\x03\x04` local file headers, ignoring the central directory entirely (which may be corrupt). Each entry is decompressed with a brute-force shift sweep (offsets 0..47) using the **ImmortalInflate** decoder ported from [socrtwo/Universal-File-Repair-Tool](https://github.com/socrtwo/Universal-File-Repair-Tool) — a never-throws DEFLATE implementation that returns whatever bytes it managed to decode before hitting bad data.
3. **DOCX XML tag fixes.** Applies the original SourceForge tool's `InvalidTags` → `ValidTags` substitutions to `word/document.xml`:
   - `mc:AlternateContent` / `mc:Choice` reordering for `wps`/`wpg`/`wpi`/`wpc`
   - `<m:oMath>` namespace repair
   - vshape / vtextbox unwrapping
   - Stripping broken `<mc:Fallback><w:pict/>` blocks
4. **RTF stripping.** Removes control words, unescapes `\'hh` and `\uN` runs, normalises `\par` / `\line` / `\tab`.
5. **Strings scan.** Last-resort scan over raw bytes for runs of printable ASCII and UTF-16LE — works on `.doc` binaries, renamed temp files, fragments dug up by `photorec`, etc.

The app also produces a **fresh, valid `.docx`** built from whatever Method 2 recovered, with a clean `[Content_Types].xml` / `_rels` and the repaired `word/document.xml` swapped in.

The **Manual recovery** tab replicates the original tool's per-OS instructions: AutoRecover paths, Previous Versions / Shadow Copies, Time Machine, Word's *Open and Repair* and *Recover Text from Any File*, and the deleted-file recovery utilities (Recuva, PhotoRec, Sleuth Kit, ShadowExplorer).

### Verified

The engine has been smoke-tested against a known-corrupt `.docx` whose `word/document.xml` deflate stream is broken (standard `unzip` fails with *invalid compressed data to inflate*). Method 2 recovered all 12 archive entries (`word/document.xml` as a partial), and text extraction produced ~5 KB of coherent document text from the recovered XML.

## Building locally

### Web (no build step)

Just serve `web/`:

```bash
cd web
python3 -m http.server 8080
# open http://localhost:8080/
```

### Desktop (Tauri 2)

Requires Rust + platform toolchains. See <https://tauri.app/v2/guides/prerequisites/>.

```bash
cd src-tauri
cargo install tauri-cli --version '^2.0.0' --locked
cargo tauri dev          # hot-reload dev window
cargo tauri build        # native installers in src-tauri/target/release/bundle/
```

### Mobile (Capacitor 6)

```bash
cd mobile
npm install
npx cap add android      # one-time
npx cap add ios          # one-time, macOS only
npx cap sync
npx cap run android      # or: open android
npx cap run ios          # or: open ios
```

See [`mobile/README.md`](mobile/README.md) for more.

## CI / releases

Pushing a tag like `v3.1.0` triggers [`.github/workflows/release.yml`](.github/workflows/release.yml) which builds installers for all six targets and attaches them to the GitHub release. The PWA at <https://socrtwo.github.io/wordrecovery-SF/> redeploys automatically from the existing [`.github/workflows/pages.yml`](.github/workflows/pages.yml) on every push to `main`.

## Repository layout

```
web/                modern web app (deployed to GitHub Pages, also wraps as PWA)
  index.html
  app.js            recovery engine
  styles.css
  manifest.webmanifest, sw.js, icons/
socrtwo/socrtwo/    mirror copy of the web app at the user-requested path
src-tauri/          Tauri 2 desktop wrapper (Win / macOS / Linux)
mobile/             Capacitor wrapper (Android / iOS)
.github/workflows/  pages deploy + multi-platform release
```

## Origin

This project was originally hosted on SourceForge as *Word Recovery* (the C# WinForms project `DocCorruptionChecker`) and migrated to GitHub. The 3.x rewrite preserves the original tool's actual repair algorithms — the `InvalidTags` / `ValidTags` substitutions for `word/document.xml`, AutoRecover-path lookups, and the link aggregation to other recovery utilities — and pairs them with the **Universal-File-Repair-Tool**'s byte-level ZIP scanner so a single web app can repair `.docx` files whose deflate streams are damaged.

- **Original SourceForge:** <https://sourceforge.net/projects/wordrecovery/>
- **Universal-File-Repair-Tool:** <https://github.com/socrtwo/Universal-File-Repair-Tool>
- **Migrated with:** [SF2GH Migrator](https://github.com/socrtwo/sf-to-github)

## Contributing

Issues and pull requests welcome at <https://github.com/socrtwo/wordrecovery-SF/issues>. The web app is plain HTML / JS — no build step — so just edit `web/app.js` and refresh.

## License

MIT — see [LICENSE](LICENSE).

---

*Maintained by [@socrtwo](https://github.com/socrtwo)*

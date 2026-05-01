# Word Recovery — Mobile

Capacitor wrapper that packages the `web/` folder as native Android and iOS apps.

## Build locally

```bash
cd mobile
npm install
npx cap add android      # one-time
npx cap add ios          # one-time, macOS only
npx cap sync             # whenever web/ changes
npx cap open android     # open Android Studio
npx cap open ios         # open Xcode
```

## Build a release APK without Android Studio

```bash
npx cap sync android
cd android
./gradlew assembleRelease
# APK at: android/app/build/outputs/apk/release/app-release-unsigned.apk
```

## Notes

- `webDir` points at `../web` so Capacitor bundles the same web build the desktop and PWA versions use.
- iOS builds need a free Apple developer team for ad-hoc / unsigned testing. For App Store distribution sign with your paid team in Xcode.
- The CI workflow at `.github/workflows/release.yml` builds an unsigned APK on every tag push.

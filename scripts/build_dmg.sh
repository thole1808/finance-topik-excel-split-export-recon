#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_ROOT"

if [[ -z "${PYTHON:-}" ]]; then
    if [[ -x "$PROJECT_ROOT/.venv/bin/python" ]]; then
        PYTHON="$PROJECT_ROOT/.venv/bin/python"
    else
        PYTHON="python3"
    fi
fi

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "Build DMG hanya bisa dijalankan di macOS." >&2
    exit 1
fi

if ! command -v hdiutil >/dev/null 2>&1; then
    echo "hdiutil tidak ditemukan; pastikan Xcode Command Line Tools tersedia." >&2
    exit 1
fi

if ! "$PYTHON" -c "import fastexcel" >/dev/null 2>&1; then
    echo "fastexcel tidak tersedia pada Python build: $PYTHON" >&2
    echo "Pasang dependency proyek dengan: $PYTHON -m pip install -r requirements.txt" >&2
    exit 1
fi

mkdir -p dist
APP_VERSION="$($PYTHON -c "from src.version import __version__; print(__version__)")"
BUILD_TIMESTAMP="$(date '+%Y%m%d_%H%M%S')"
BUILD_DIR="$PROJECT_ROOT/build/macos-${APP_VERSION}-${BUILD_TIMESTAMP}"
APP_BUNDLE="$BUILD_DIR/Recon Builder.app"

# ── 1. PyInstaller ─────────────────────────────────────────────────────────
"$PYTHON" -m PyInstaller \
    --noconfirm \
    --windowed \
    --name "Recon Builder" \
    --icon "$PROJECT_ROOT/assets/recon-builder.icns" \
    --collect-all fastexcel \
    --hidden-import fastexcel \
    --collect-all PIL \
    --hidden-import PIL \
    --collect-all certifi \
    --hidden-import certifi \
    --paths "$PROJECT_ROOT" \
    --add-data "$PROJECT_ROOT/src/ref_styles.json:src" \
    --add-data "$PROJECT_ROOT/src/fee_tables.json:src" \
    --add-data "$PROJECT_ROOT/src/reference_data.json:src" \
    --add-data "$PROJECT_ROOT/assets/batik-watermark.png:assets" \
    --distpath "$BUILD_DIR" \
    --specpath "$PROJECT_ROOT/build" \
    --workpath "$PROJECT_ROOT/build/pyinstaller" \
    "$PROJECT_ROOT/app.py"

DMG_NAME="ReconBuilder-v${APP_VERSION}-${BUILD_TIMESTAMP}.dmg"
PKG_NAME="ReconBuilder-v${APP_VERSION}-${BUILD_TIMESTAMP}.pkg"

# ── 2. Patch Info.plist ────────────────────────────────────────────────────
/usr/libexec/PlistBuddy \
    -c "Set :CFBundleShortVersionString $APP_VERSION" \
    -c "Set :CFBundleName Recon Builder" \
    -c "Set :CFBundleDisplayName Recon Builder" \
    -c "Set :CFBundleIdentifier com.devops.reconbuilder" \
    "$APP_BUNDLE/Contents/Info.plist"
# Add CFBundleVersion hanya jika belum ada
/usr/libexec/PlistBuddy -c "Add :CFBundleVersion string $APP_VERSION" \
    "$APP_BUNDLE/Contents/Info.plist" 2>/dev/null || \
/usr/libexec/PlistBuddy -c "Set :CFBundleVersion $APP_VERSION" \
    "$APP_BUNDLE/Contents/Info.plist"

codesign --force --deep --sign - "$APP_BUNDLE"

# ── 3. Hapus artefak lama ──────────────────────────────────────────────────
for f in \
    "$PROJECT_ROOT"/ReconBuilder-v*.dmg \
    "$PROJECT_ROOT"/ReconBuilder-v*.pkg \
    "$PROJECT_ROOT/dist"/ReconBuilder-v*.dmg \
    "$PROJECT_ROOT/dist"/ReconBuilder-v*.pkg; do
    [[ -f "$f" ]] && rm -f "$f"
done

# ── 4. DMG (ULFO = lzfse, jauh lebih cepat dari UDZO di Apple runner) ─────
DMG_STAGE="$PROJECT_ROOT/dist/macos-installer-stage"
rm -rf "$DMG_STAGE"
mkdir -p "$DMG_STAGE"
ditto --rsrc --extattr "$APP_BUNDLE" "$DMG_STAGE/Recon Builder.app"
ln -s /Applications "$DMG_STAGE/Applications"
printf 'Buka DMG, lalu seret Recon Builder.app ke Applications. Selesai.\n' \
    > "$DMG_STAGE/CARA INSTALL.txt"

hdiutil create \
    -ov \
    -format ULFO \
    -srcowners off \
    -volname "Recon Builder" \
    -srcfolder "$DMG_STAGE" \
    "$PROJECT_ROOT/dist/$DMG_NAME"
rm -rf "$DMG_STAGE"

# ── 5. PKG (pkgbuild langsung tanpa --analyze pass) ───────────────────────
PKG_ROOT="$BUILD_DIR/pkgroot"
mkdir -p "$PKG_ROOT/Applications"
# Reuse bundle yang sudah ada di BUILD_DIR, tidak perlu ditto ulang
cp -a "$APP_BUNDLE" "$PKG_ROOT/Applications/Recon Builder.app"

pkgbuild \
    --root "$PKG_ROOT" \
    --scripts "$PROJECT_ROOT/scripts/macos" \
    --install-location / \
    --identifier com.devops.reconbuilder \
    --version "$APP_VERSION" \
    "$PROJECT_ROOT/dist/$PKG_NAME"

# Salin ke root untuk backward compat
cp "$PROJECT_ROOT/dist/$DMG_NAME" "$PROJECT_ROOT/$DMG_NAME"
cp "$PROJECT_ROOT/dist/$PKG_NAME" "$PROJECT_ROOT/$PKG_NAME"
rm -f "$PROJECT_ROOT/ReconBuilder.dmg"
echo "DMG berhasil dibuat: $PROJECT_ROOT/dist/$DMG_NAME"

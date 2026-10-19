#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_ROOT"

CLEAN=0
while [[ $# -gt 0 ]]; do
    case "$1" in
        --clean)
            CLEAN=1
            shift
            ;;
        -h|--help)
            echo "Penggunaan: build_dmg.sh [--clean]"
            echo "  --clean  Hapus folder build/ dan dist/ sebelum build"
            exit 0
            ;;
        *)
            echo "Argumen tidak dikenal: $1" >&2
            exit 1
            ;;
    esac
done

if [[ -z "${PYTHON:-}" ]]; then
    if [[ -x "$PROJECT_ROOT/.venv/bin/python" ]]; then
        PYTHON="$PROJECT_ROOT/.venv/bin/python"
    else
        PYTHON="python3"
    fi
fi

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "Build DMG/PKG hanya bisa dijalankan di macOS." >&2
    exit 1
fi

if ! command -v hdiutil >/dev/null 2>&1; then
    echo "hdiutil tidak ditemukan; pastikan Xcode Command Line Tools terpasang." >&2
    echo "Instal dengan: xcode-select --install" >&2
    exit 1
fi

if ! command -v pkgbuild >/dev/null 2>&1; then
    echo "pkgbuild tidak ditemukan; pastikan Xcode Command Line Tools terpasang." >&2
    exit 1
fi

if ! command -v codesign >/dev/null 2>&1; then
    echo "codesign tidak ditemukan; pastikan Xcode Command Line Tools terpasang." >&2
    exit 1
fi

echo "Memverifikasi dependensi Python..."
for module in fastexcel polars xlsxwriter certifi; do
    if ! "$PYTHON" -c "import $module" >/dev/null 2>&1; then
        echo "Module Python '$module' tidak tersedia pada $PYTHON." >&2
        echo "Pasang dependency dengan: $PYTHON -m pip install -r requirements.txt" >&2
        exit 1
    fi
done

# Bersihkan build lama jika diminta
if [[ "$CLEAN" -eq 1 ]]; then
    echo "Membersihkan build/ dan dist/..."
    chmod -R +w "$PROJECT_ROOT/build" "$PROJECT_ROOT/dist" 2>/dev/null || true
    rm -rf "$PROJECT_ROOT/build" "$PROJECT_ROOT/dist" 2>/dev/null || true
fi

mkdir -p "$PROJECT_ROOT/dist"

# Keep temporary app/build products under dist and clear leftovers from any
# interrupted packaging run before starting a fresh build.
rm -rf "$PROJECT_ROOT/dist/.build" \
    "$PROJECT_ROOT/dist/macos-installer-stage"
mkdir -p "$PROJECT_ROOT/dist/.build/spec" \
    "$PROJECT_ROOT/dist/.build/pyinstaller"

APP_VERSION="$($PYTHON -c "from src.version import __version__; print(__version__)")"
BUILD_TIMESTAMP="$(date '+%Y%m%d_%H%M%S')"
BUILD_DIR="$PROJECT_ROOT/dist/.build/macos-${APP_VERSION}-${BUILD_TIMESTAMP}"
APP_BUNDLE="$BUILD_DIR/Auto Recon Biller.app"

echo "=== Build Auto Recon Biller v$APP_VERSION ($BUILD_TIMESTAMP) ==="
echo "Python: $PYTHON"
echo "Build dir: $BUILD_DIR"

# ── 1. PyInstaller ─────────────────────────────────────────────────────────
echo "Membangun aplikasi dengan PyInstaller..."
"$PYTHON" -m PyInstaller \
    --noconfirm \
    --windowed \
    --name "Auto Recon Biller" \
    --icon "$PROJECT_ROOT/assets/recon-builder.icns" \
    --collect-all fastexcel \
    --hidden-import fastexcel \
    --collect-all certifi \
    --hidden-import certifi \
    --paths "$PROJECT_ROOT" \
    --add-data "$PROJECT_ROOT/src/ref_styles.json:src" \
    --add-data "$PROJECT_ROOT/src/fee_tables.json:src" \
    --add-data "$PROJECT_ROOT/src/reference_data.json:src" \
    --distpath "$BUILD_DIR" \
    --specpath "$PROJECT_ROOT/dist/.build/spec" \
    --workpath "$PROJECT_ROOT/dist/.build/pyinstaller" \
    "$PROJECT_ROOT/app.py"

DMG_NAME="ReconBuilder-v${APP_VERSION}-${BUILD_TIMESTAMP}.dmg"
PKG_NAME="ReconBuilder-v${APP_VERSION}-${BUILD_TIMESTAMP}.pkg"

# ── 2. Patch Info.plist ────────────────────────────────────────────────────
echo "Mengkonfigurasi Info.plist..."
/usr/libexec/PlistBuddy \
    -c "Set :CFBundleShortVersionString $APP_VERSION" \
    -c "Set :CFBundleName Auto Recon Biller" \
    -c "Set :CFBundleDisplayName Auto Recon Biller" \
    -c "Set :CFBundleIdentifier com.devops.reconbuilder" \
    "$APP_BUNDLE/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Add :CFBundleVersion string $APP_VERSION" \
    "$APP_BUNDLE/Contents/Info.plist" 2>/dev/null || \
/usr/libexec/PlistBuddy -c "Set :CFBundleVersion $APP_VERSION" \
    "$APP_BUNDLE/Contents/Info.plist"

codesign --force --deep --sign - "$APP_BUNDLE"

# ── 3. Hapus artefak lama ──────────────────────────────────────────────────
echo "Membersihkan artefak lama..."
for f in \
    "$PROJECT_ROOT"/ReconBuilder-v*.dmg \
    "$PROJECT_ROOT"/ReconBuilder-v*.pkg \
    "$PROJECT_ROOT/dist"/ReconBuilder-v*.dmg \
    "$PROJECT_ROOT/dist"/ReconBuilder-v*.pkg; do
    [[ -f "$f" ]] && rm -f "$f"
done
rm -f "$PROJECT_ROOT/ReconBuilder.dmg" "$PROJECT_ROOT/ReconBuilder.pkg"
rm -f "$PROJECT_ROOT"/ReconBuilder-v*.dmg.zip

# ── 4. DMG (ULFO = lzfse, jauh lebih cepat dari UDZO di Apple runner) ─────
echo "Membuat DMG..."
DMG_STAGE="$PROJECT_ROOT/dist/macos-installer-stage"
rm -rf "$DMG_STAGE"
mkdir -p "$DMG_STAGE"
ditto --rsrc --extattr "$APP_BUNDLE" "$DMG_STAGE/Auto Recon Biller.app"
ln -s /Applications "$DMG_STAGE/Applications"
printf 'Buka DMG, lalu seret Auto Recon Biller.app ke Applications. Selesai.\n' \
    > "$DMG_STAGE/CARA INSTALL.txt"

hdiutil create \
    -ov \
    -format ULFO \
    -srcowners off \
    -volname "Auto Recon Biller" \
    -srcfolder "$DMG_STAGE" \
    "$PROJECT_ROOT/dist/$DMG_NAME"
rm -rf "$DMG_STAGE"

# ── 5. PKG (pkgbuild langsung tanpa --analyze pass) ───────────────────────
echo "Membuat PKG..."
PKG_ROOT="$BUILD_DIR/pkgroot"
mkdir -p "$PKG_ROOT/Applications"
cp -a "$APP_BUNDLE" "$PKG_ROOT/Applications/Auto Recon Biller.app"

pkgbuild \
    --root "$PKG_ROOT" \
    --scripts "$PROJECT_ROOT/scripts/macos" \
    --install-location / \
    --identifier com.devops.reconbuilder \
    --version "$APP_VERSION" \
    "$PROJECT_ROOT/dist/$PKG_NAME"

# Keep only the final installers in dist after packaging completes.
rm -rf "$PROJECT_ROOT/dist/.build"

echo ""
echo "=== Build selesai ==="
echo "DMG:  $PROJECT_ROOT/dist/$DMG_NAME"
echo "PKG:  $PROJECT_ROOT/dist/$PKG_NAME"
echo "Semua file build tersedia di dist/."

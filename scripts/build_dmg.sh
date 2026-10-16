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
"$PYTHON" -m PyInstaller \
    --clean \
    --noconfirm \
    --windowed \
    --name "ReconBuilder" \
    --collect-all fastexcel \
    --hidden-import fastexcel \
    --paths "$PROJECT_ROOT" \
    --add-data "$PROJECT_ROOT/src/ref_styles.json:src" \
    --add-data "$PROJECT_ROOT/src/fee_tables.json:src" \
    --add-data "$PROJECT_ROOT/src/reference_data.json:src" \
    --distpath "$PROJECT_ROOT/dist" \
    --specpath "$PROJECT_ROOT/build" \
    --workpath "$PROJECT_ROOT/build/pyinstaller" \
    "$PROJECT_ROOT/app.py"

APP_VERSION="$("$PYTHON" -c "from src.version import __version__; print(__version__)")"
/usr/libexec/PlistBuddy \
    -c "Set :CFBundleShortVersionString $APP_VERSION" \
    "$PROJECT_ROOT/dist/ReconBuilder.app/Contents/Info.plist"
/usr/libexec/PlistBuddy \
    -c "Add :CFBundleVersion string $APP_VERSION" \
    "$PROJECT_ROOT/dist/ReconBuilder.app/Contents/Info.plist"

hdiutil create \
    -ov \
    -format UDZO \
    -volname "Recon Builder" \
    -srcfolder "$PROJECT_ROOT/dist/ReconBuilder.app" \
    "$PROJECT_ROOT/dist/ReconBuilder.dmg"

echo "DMG berhasil dibuat: $PROJECT_ROOT/dist/ReconBuilder.dmg"

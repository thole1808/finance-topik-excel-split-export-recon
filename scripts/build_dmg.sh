#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_ROOT"

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "Build DMG hanya bisa dijalankan di macOS." >&2
    exit 1
fi

if ! command -v hdiutil >/dev/null 2>&1; then
    echo "hdiutil tidak ditemukan; pastikan Xcode Command Line Tools tersedia." >&2
    exit 1
fi

mkdir -p dist
python3 -m PyInstaller \
    --clean \
    --noconfirm \
    --windowed \
    --name "ReconBuilder" \
    --paths "$PROJECT_ROOT" \
    --add-data "$PROJECT_ROOT/src/ref_styles.json:src" \
    --add-data "$PROJECT_ROOT/src/fee_tables.json:src" \
    --add-data "$PROJECT_ROOT/src/reference_data.json:src" \
    --distpath "$PROJECT_ROOT/dist" \
    --specpath "$PROJECT_ROOT/build" \
    --workpath "$PROJECT_ROOT/build/pyinstaller" \
    "$PROJECT_ROOT/app.py"

hdiutil create \
    -ov \
    -format UDZO \
    -volname "Recon Builder" \
    -srcfolder "$PROJECT_ROOT/dist/ReconBuilder.app" \
    "$PROJECT_ROOT/dist/ReconBuilder.dmg"

echo "DMG berhasil dibuat: $PROJECT_ROOT/dist/ReconBuilder.dmg"

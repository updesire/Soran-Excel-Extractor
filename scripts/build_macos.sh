#!/bin/bash
set -euo pipefail

ROOT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)"
DIST_DIR="$ROOT_DIR/dist"

APP_NAME="${APP_NAME:-SoranExcelExtractor}"
BUNDLE_ID="${BUNDLE_ID:-com.soran.excel.extractor}"
VERSION="${VERSION:-1.0.0}"

APP_DIR="$DIST_DIR/$APP_NAME.app"
CONTENTS_DIR="$APP_DIR/Contents"
MACOS_DIR="$CONTENTS_DIR/MacOS"
RESOURCES_DIR="$CONTENTS_DIR/Resources"

PYTHON_SCRIPT_SRC="$ROOT_DIR/process_excel_files.py"

mkdir -p "$DIST_DIR"
rm -rf "$APP_DIR"
mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

cat > "$CONTENTS_DIR/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleDevelopmentRegion</key>
  <string>en</string>
  <key>CFBundleDisplayName</key>
  <string>$APP_NAME</string>
  <key>CFBundleExecutable</key>
  <string>$APP_NAME</string>
  <key>CFBundleIdentifier</key>
  <string>$BUNDLE_ID</string>
  <key>CFBundleInfoDictionaryVersion</key>
  <string>6.0</string>
  <key>CFBundleName</key>
  <string>$APP_NAME</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleShortVersionString</key>
  <string>$VERSION</string>
  <key>CFBundleVersion</key>
  <string>$VERSION</string>
  <key>LSMinimumSystemVersion</key>
  <string>10.13</string>
  <key>NSHighResolutionCapable</key>
  <true/>
</dict>
</plist>
EOF

cat > "$MACOS_DIR/$APP_NAME" <<'EOF'
#!/bin/sh
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
RESOURCES_DIR="$(CDPATH= cd -- "$SCRIPT_DIR/../Resources" && pwd)"
RUN_CMD="$RESOURCES_DIR/run_app.command"
exec /usr/bin/open "$RUN_CMD"
EOF

chmod +x "$MACOS_DIR/$APP_NAME"

cp "$PYTHON_SCRIPT_SRC" "$RESOURCES_DIR/process_excel_files.py"

cat > "$RESOURCES_DIR/run_app.command" <<'EOF'
#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
APP_BUNDLE_DIR="$(CDPATH= cd -- "$SCRIPT_DIR/../.." && pwd)"
BASE_DIR="$(CDPATH= cd -- "$(dirname -- "$APP_BUNDLE_DIR")" && pwd)"

INPUT_DIR="$BASE_DIR/source"
OUTPUT_DIR="$BASE_DIR/output"
PYTHON_BIN="/usr/bin/python3"
PYTHON_SCRIPT="$SCRIPT_DIR/process_excel_files.py"

mkdir -p "$INPUT_DIR" "$OUTPUT_DIR"

clear
printf "\n========================================\n"
printf "        SORAN EXCEL EXTRACTION TOOL\n"
printf "========================================\n\n"
printf "Input : %s\n" "$INPUT_DIR"
printf "Output: %s\n\n" "$OUTPUT_DIR"

"$PYTHON_BIN" -u "$PYTHON_SCRIPT" --input "$INPUT_DIR" --output "$OUTPUT_DIR"
exit_code=$?

printf "\n----------------------------------------\n"
if [ "$exit_code" -eq 0 ]; then
  printf "Finished successfully.\n"
else
  printf "Failed with exit code %s.\n" "$exit_code"
fi
printf "----------------------------------------\n"
printf "Press Enter to close this window..."
read -r _
EOF

chmod +x "$RESOURCES_DIR/run_app.command"

/usr/bin/plutil -lint "$CONTENTS_DIR/Info.plist" >/dev/null

/usr/bin/codesign --force --deep --sign - "$APP_DIR" || true

DMG_PATH="$DIST_DIR/$APP_NAME.dmg"
rm -f "$DMG_PATH"
/usr/bin/hdiutil create -volname "$APP_NAME" -srcfolder "$APP_DIR" -ov -format UDZO "$DMG_PATH" >/dev/null

printf "%s\n" "$APP_DIR"
printf "%s\n" "$DMG_PATH"

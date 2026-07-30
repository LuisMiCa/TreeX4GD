#!/usr/bin/env bash

echo "----------------------------------------"
echo "TreeX4GD — Single-file macOS App Builder"
echo "----------------------------------------"
# Produces a TreeX4GD.app bundle using PyInstaller

# ------------------------------------------------------------
# 0. OS Validation
# ------------------------------------------------------------
OS_NAME=$(uname -s)
if [ "$OS_NAME" = "Linux" ]; then
    echo "❌ This script is for macOS. Please run the Linux version."
    read -p "Press Enter to exit..." dummy
    exit 1
elif [ "$OS_NAME" != "Darwin" ]; then
    echo "⚠️  Warning: Unknown OS ($OS_NAME). Proceeding anyway..."
fi

# ------------------------------------------------------------
# 1. Resolve paths
# ------------------------------------------------------------
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
APP_SCRIPT_DIR="$PROJECT_ROOT/script"
TMP_DIR="$SCRIPT_DIR/TMP"
FINAL_DIST_DIR="$SCRIPT_DIR/TreeX4GD-build-output"
REQUIREMENTS="$APP_SCRIPT_DIR/requirements.txt"

echo "Project root: $PROJECT_ROOT"
echo ""

# ------------------------------------------------------------
# 2. TRAP: Automatic Cleanup & Pause on Exit
# ------------------------------------------------------------
cleanup() {
    EXIT_CODE=$?
    # 2.1. Alert if it crashed unexpectedly (Ignore 0 for success, and 1 for handled errors)
    if [ $EXIT_CODE -ne 0 ] && [ $EXIT_CODE -ne 1 ]; then
        echo ""
        echo "⚠️ Unexpected crash! Script exited with error code: $EXIT_CODE"
    fi

    # 2.2. Cleanup temp files if they exist
    if [ -d "$TMP_DIR" ]; then
        echo ""
        echo "🧹 Cleaning up temporary build directory..."
        rm -rf "$TMP_DIR"
    fi

    # 2.3. Pause before closing if running in a terminal (useful if was double-clicked)
    echo ""
    if [ -t 0 ]; then
        read -p "You can close this window now..." dummy
    fi
}
trap cleanup EXIT

# ------------------------------------------------------------
# 3. Pre-Flight Path Check
# ------------------------------------------------------------
echo "🔍 Checking project path for problematic characters..."
if [[ "$PROJECT_ROOT" == *"["* ]] || [[ "$PROJECT_ROOT" == *"]"* ]]; then
    echo "❌ [FATAL] INVALID FOLDER NAME DETECTED"
    echo "PyInstaller may incorrectly parse the project path if it"
    echo "contains certain special characters, such as square brackets [ ]."
    echo "Current Path: $PROJECT_ROOT"
    echo "Please rename your project folder."
    exit 1
fi
echo "✅ Path is safe for PyInstaller."
echo ""

# ------------------------------------------------------------
# 4. Check Xcode Command Line Tools
# ------------------------------------------------------------
# On macOS, PyInstaller relies on Apple tools like 'lipo' to build universal binaries.
echo "🔍 Checking Xcode Command Line Tools…"

if ! xcode-select -p >/dev/null 2>&1; then
    echo ""
    echo "❌ Xcode Command Line Tools are not installed."
    echo "PyInstaller requires Apple tooling to build macOS apps."
    echo ""
    echo "📦 Triggering installer…"
    echo "➡️  Complete the tools' installation and then re-run this script."
    echo ""
    xcode-select --install
    exit 0   # Exits cleanly so the trap doesn't throw a generic failure message
fi

if ! command -v lipo >/dev/null 2>&1; then
    echo ""
    echo "❌ 'lipo' not found even though CLT appears installed."
    echo "Try reinstalling Command Line Tools."
    exit 1
fi

echo "✅ Xcode Command Line Tools detected."
echo ""

# ------------------------------------------------------------
# 5. Virtual Environment & Dependency Validation
# ------------------------------------------------------------
echo "🔍 Searching for Python virtual environment…"
VENV_DIR="$APP_SCRIPT_DIR/venv"
VENV_PYTHON="$VENV_DIR/bin/python"

# 5.1. Check if virtual environment exists
if [ ! -x "$VENV_PYTHON" ]; then
    echo "❌ No virtual environment found at:"
    echo "  $VENV_PYTHON"
    echo "👉 Please run the script 'run_in_macOS.command' first to prepare the environment."
    exit 1
fi

# 5.2. Python Version Check (Ensures Python 3.11+)
if ! "$VENV_PYTHON" -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >/dev/null 2>&1; then
    echo "❌ Virtual environment Python version is too old (requires 3.11+)."
    exit 1
fi

# 5.3. Requirements File check
if [ ! -f "$REQUIREMENTS" ]; then
    echo "❌ Requirements file missing at: $REQUIREMENTS"
    exit 1
fi

# 5.4. Check if all dependencies are installed in the venv
"$VENV_PYTHON" -c "
import sys, re; from importlib.metadata import distributions
installed = {d.metadata['Name'].lower().replace('-', '_') for d in distributions()}
missing = any(
    (m := re.match(r'^([a-zA-Z0-9_\-\.]+)', l.split('#')[0].strip())) and 
    m.group(1).lower().replace('-', '_') not in installed
    for l in open(\"$REQUIREMENTS\")
)
sys.exit(missing)
" >/dev/null 2>&1

if [ $? -ne 0 ]; then
    echo "❌ Dependencies are missing from the virtual environment."
    echo "👉 Please run the script 'run_in_macOS.command' first to install requirements."
    exit 1
fi
echo "✅ Virtual environment and dependencies are healthy."
echo ""

# ------------------------------------------------------------
# 6. PyInstaller & Internet Check
# ------------------------------------------------------------
echo "🔍 Checking PyInstaller…"
if ! "$VENV_PYTHON" -m PyInstaller --version >/dev/null 2>&1; then
    echo "📦 PyInstaller not found. Installation required."
    echo "🔍 Verifying internet connection..."
    
    # Using socket to verify outward connection without triggering macOS SSL cert errors
    if ! "$VENV_PYTHON" -c "import socket; socket.create_connection(('www.google.com', 443), timeout=3)" > /dev/null 2>&1; then
        echo "❌ You are OFFLINE."
        echo "PyInstaller needs to be downloaded to build the executable,"
        echo "but no internet connection was detected."
        exit 1
    fi
    
    echo "✅ Internet connected. Installing PyInstaller..."
    "$VENV_PYTHON" -m pip install --upgrade pip --quiet
    "$VENV_PYTHON" -m pip install pyinstaller
    if [ $? -ne 0 ]; then
        echo "❌ Failed to install PyInstaller."
        exit 1
    fi
    echo "✅ PyInstaller installed successfully."
else
    echo "✅ PyInstaller already installed."
fi
echo ""

# ------------------------------------------------------------
# 7. Pre-Flight Check & Asset Verification
# ------------------------------------------------------------
echo "📦 Verifying source assets…"
MISSING_FILES=false

SRC_SPEC="$SCRIPT_DIR/TreeX4GD.spec"
SRC_ICON="$SCRIPT_DIR/TreeX4GD.icns"
SRC_PYTHON="$APP_SCRIPT_DIR/TreeX4GD.py"
SRC_TPL_HTML="$APP_SCRIPT_DIR/TreeX4GD-template.html"
SRC_HELP_HTML="$APP_SCRIPT_DIR/TreeX4GD-help.html"
SRC_IMAGES_DIR="$APP_SCRIPT_DIR/TreeX4GD-help-images"

[ ! -f "$SRC_SPEC" ] && MISSING_FILES=true && echo "❌ Missing: $SRC_SPEC"
[ ! -f "$SRC_ICON" ] && MISSING_FILES=true && echo "❌ Missing: $SRC_ICON"
[ ! -f "$SRC_PYTHON" ] && MISSING_FILES=true && echo "❌ Missing: $SRC_PYTHON"
[ ! -f "$SRC_TPL_HTML" ] && MISSING_FILES=true && echo "❌ Missing: $SRC_TPL_HTML"
[ ! -f "$SRC_HELP_HTML" ] && MISSING_FILES=true && echo "❌ Missing: $SRC_HELP_HTML"
[ ! -d "$SRC_IMAGES_DIR" ] && MISSING_FILES=true && echo "❌ Missing Folder: $SRC_IMAGES_DIR"

if [ "$MISSING_FILES" = true ]; then
    echo "❌ [FATAL] Required build files are missing."
    exit 1
fi
echo "✅ All source files found."
echo ""

# ------------------------------------------------------------
# 8. Prepare workspace and copy files
# ------------------------------------------------------------
echo "📄 Copying build files to temporary build directory…"
rm -rf "$TMP_DIR"
mkdir "$TMP_DIR"

cp "$SRC_SPEC" "$TMP_DIR/"
cp "$SRC_ICON" "$TMP_DIR/"
cp "$SRC_PYTHON" "$TMP_DIR/"
cp "$SRC_TPL_HTML" "$TMP_DIR/"
cp "$SRC_HELP_HTML" "$TMP_DIR/"
cp -R "$SRC_IMAGES_DIR" "$TMP_DIR/"
echo "✅ Files copied to $TMP_DIR"
echo ""

# ------------------------------------------------------------
# 9. Build macOS app bundle
# ------------------------------------------------------------
cd "$TMP_DIR"

echo "🚀 Running PyInstaller…"
"$VENV_PYTHON" -m PyInstaller --clean TreeX4GD.spec

if [ $? -ne 0 ]; then
    echo "❌ Build failed. See PyInstaller output above."
    exit 1
fi

echo ""
echo "📦 Build finished — processing output…"

# ------------------------------------------------------------
# 10. Move final dist output
# ------------------------------------------------------------
TMP_DIST="$TMP_DIR/dist"
APP_PATH="$TMP_DIST/TreeX4GD.app"

if [ ! -d "$APP_PATH" ]; then
    echo "❌ Build failed: TreeX4GD.app not found."
    echo "Contents of dist:"
    ls -l "$TMP_DIST"
    exit 1
fi

echo "➡️  Moving TreeX4GD.app to:"
echo "   $FINAL_DIST_DIR"

rm -rf "$FINAL_DIST_DIR"
mkdir -p "$FINAL_DIST_DIR"
mv "$APP_PATH" "$FINAL_DIST_DIR/"

# ------------------------------------------------------------
# 11. Apple code signing (for Apple Silicon compatibility)
# ------------------------------------------------------------
# This ad-hoc signature prevents "App Damaged" errors on ARM Macs
echo "🔏 Applying ad-hoc code signature..."
codesign --force --deep --sign - "$FINAL_DIST_DIR/TreeX4GD.app"
if [ $? -ne 0 ]; then
    echo "❌ Code signing failed."
    exit 1
fi
echo "✅ App signed (Ad-Hoc)."

# ------------------------------------------------------------
# 12. Final Notes
# ------------------------------------------------------------
echo ""
echo "✅ Build complete!"
echo ""
echo "🚀 App bundle location:"
echo "  $FINAL_DIST_DIR/TreeX4GD.app"
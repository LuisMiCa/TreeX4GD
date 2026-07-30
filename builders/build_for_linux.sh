#!/usr/bin/env bash

echo "----------------------------------------"
echo "TreeX4GD — Single-file Linux App Builder"
echo "----------------------------------------"
# Produces a standalone executable binary using PyInstaller

# ------------------------------------------------------------
# 0. OS Validation
# ------------------------------------------------------------
OS_NAME=$(uname -s)
if [ "$OS_NAME" = "Darwin" ]; then
    echo "❌ This script is for Linux. Please run the macOS version."
    read -p "Press Enter to exit..."
    exit 1
elif [ "$OS_NAME" != "Linux" ]; then
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
    echo ""
    
    # 2.3. Pause before closing if running in a terminal (useful if was double-clicked)
    if [ -t 0 ]; then
        read -p "Press Enter to close..."
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
# 4. Check System Dependencies (objdump)
# ------------------------------------------------------------
# PyInstaller on Linux requires 'objdump' (from binutils) to analyze shared libraries.
echo "🔍 Checking system build tools..."
if ! command -v objdump >/dev/null 2>&1; then
    echo "❌ 'objdump' is not installed."
    echo "PyInstaller requires this tool to analyze binary dependencies."
    echo "💡 To fix this, please install 'binutils' via your package manager."
    echo "   - Ubuntu/Debian: sudo apt install binutils"
    echo "   - Fedora/RHEL:   sudo dnf install binutils"
    echo "   - Arch/Manjaro:  sudo pacman -S binutils"
    echo ""
    exit 1
fi
echo "✅ 'objdump' detected."
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
    echo "👉 Please run the script 'run_in_linux.sh' first to prepare the environment."
    exit 1
fi

# 5.2. Python Version Check (Ensures Python 3.11+)
if ! "$VENV_PYTHON" -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >/dev/null 2>&1; then
    echo "❌ Virtual environment Python version is too old (requires 3.11+)."
    exit 1
fi

# 5.3. Requirements file Check
if [ ! -f "$REQUIREMENTS" ]; then
    echo "❌ Requirements file missing at: $REQUIREMENTS"
    exit 1
fi

# 5.4. Check if all requirements are installed in the venv
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

# If the above command exits with code 1, it means some dependencies are missing.
if [ $? -ne 0 ]; then
    echo "❌ Dependencies are missing from the virtual environment."
    echo "👉 Please run the script 'run_in_linux.sh' first to install requirements."
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
    
    # Check for internet connectivity before attempting to install PyInstaller
    if ! "$VENV_PYTHON" -c "import socket; socket.create_connection(('www.google.com', 443), timeout=3)" > /dev/null 2>&1; then
        echo "❌ You are OFFLINE."
        echo "PyInstaller needs to be downloaded to build the executable,"
        echo "but no internet connection was detected."
        exit 1
    fi
    
    # If we reach this point, the internet is working, so we can attempt to install PyInstaller.
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
# 7. Verify Source Files
# ------------------------------------------------------------
echo "📦 Verifying source assets…"
MISSING_FILES=false

SRC_SPEC="$SCRIPT_DIR/TreeX4GD.spec"
SRC_ICON="$APP_SCRIPT_DIR/TreeX4GD.png"
SRC_PYTHON="$APP_SCRIPT_DIR/TreeX4GD.py"
SRC_TPL_HTML="$APP_SCRIPT_DIR/TreeX4GD-template.html"
SRC_HELP_HTML="$APP_SCRIPT_DIR/TreeX4GD-help.html"
SRC_IMAGES_DIR="$APP_SCRIPT_DIR/TreeX4GD-help-images"

# Check for the presence of all required source files and directories
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
rm -rf "$TMP_DIR"
mkdir "$TMP_DIR"

echo "📄 Copying build files to temporary build directory.…"
cp "$SRC_SPEC" "$TMP_DIR/"
cp "$SRC_ICON" "$TMP_DIR/"
cp "$SRC_PYTHON" "$TMP_DIR/"
cp "$SRC_TPL_HTML" "$TMP_DIR/"
cp "$SRC_HELP_HTML" "$TMP_DIR/"
cp -R "$SRC_IMAGES_DIR" "$TMP_DIR/"
echo "✅ Files copied to $TMP_DIR"
echo ""

# ------------------------------------------------------------
# 9. Build Linux Binary
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
APP_BINARY="$TMP_DIR/dist/TreeX4GD" 

if [ ! -f "$APP_BINARY" ]; then
    echo "❌ Build failed: TreeX4GD binary not found in dist."
    exit 1
fi

echo "➡️  Moving TreeX4GD binary to: $FINAL_DIST_DIR"
rm -rf "$FINAL_DIST_DIR"
mkdir -p "$FINAL_DIST_DIR"

mv "$APP_BINARY" "$FINAL_DIST_DIR/"
chmod +x "$FINAL_DIST_DIR/TreeX4GD"

echo "✅ Binary successfully built and moved."
echo ""

# ------------------------------------------------------------
# 11. Optional Installer & .desktop Generation
# ------------------------------------------------------------
echo "Do you want to generate a Linux installer/uninstaller and .desktop launcher?"
echo "This allows users to integrate the app into their system menus (Ubuntu, Mint, Fedora, etc.)."
read -p "Generate installer files? [y/N]: " -n 1 -r
echo ""
echo ""

if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "⚙️  Generating installation files..."
    
    # 1. Copy the icon to the final dist folder
    cp "$SRC_ICON" "$FINAL_DIST_DIR/TreeX4GD.png"
    
    # 2. Generate the .desktop file dynamically
    #    and yes, it's TreeX4GD in the StartupWMClass because that's how it appears in xprop (case-sensitive)
    cat << 'EOF' > "$FINAL_DIST_DIR/TreeX4GD.desktop"
[Desktop Entry]
Version=1.0
Type=Application
Name=TreeX4GD
Comment=File Tree Exporter for Google Drive
Exec=TreeX4GD
Icon=TreeX4GD
Terminal=false
Categories=Utility;Office;
StartupWMClass=TreeX4GD
EOF

    # 3. Generate the install.sh script dynamically
    # Note: 'EOF' is quoted so variables like $HOME are evaluated during INSTALL, not during BUILD.
    cat << 'EOF' > "$FINAL_DIST_DIR/install.sh"
#!/usr/bin/env bash

echo "📦 Installing TreeX4GD for the current user..."

# Standard FreeDesktop.org paths (Compatible with all major distros)
BIN_DIR="$HOME/.local/bin"
APP_DIR="$HOME/.local/share/applications"
ICON_DIR="$HOME/.local/share/icons/hicolor/256x256/apps"

# Create directories if they don't exist
mkdir -p "$BIN_DIR" "$APP_DIR" "$ICON_DIR"

# Copy files into system locations
cp "TreeX4GD" "$BIN_DIR/"
cp "TreeX4GD.desktop" "$APP_DIR/"
cp "TreeX4GD.png" "$ICON_DIR/TreeX4GD.png"

# Ensure binary is executable
chmod +x "$BIN_DIR/TreeX4GD"

echo "🔄 Refreshing system application caches..."

# Refresh caches so the app appears immediately
if command -v update-desktop-database >/dev/null 2>&1; then update-desktop-database "$APP_DIR"; fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then gtk-update-icon-cache -f -t "$HOME/.local/share/icons/hicolor" >/dev/null 2>&1; fi
if command -v xdg-desktop-menu >/dev/null 2>&1; then xdg-desktop-menu forceupdate --mode user; fi
touch "$APP_DIR"

echo "✅ Installation complete!"
echo "--------------------------------------------------------"
echo "You can now find TreeX4GD in your application menu."
echo "Note: If it doesn't launch, ensure that '$BIN_DIR' is"
echo "added to your system's PATH environment variable."
echo "--------------------------------------------------------"
echo ""
read -p "Press Enter to close..."
EOF

    # 4. Generate the uninstall.sh script dynamically
    cat << 'EOF' > "$FINAL_DIST_DIR/uninstall.sh"
#!/usr/bin/env bash

echo "🗑️  Uninstalling TreeX4GD..."

# Standard FreeDesktop.org paths
BIN_DIR="$HOME/.local/bin"
APP_DIR="$HOME/.local/share/applications"
ICON_DIR="$HOME/.local/share/icons/hicolor/256x256/apps"

# Remove the files
rm -f "$BIN_DIR/TreeX4GD"
rm -f "$APP_DIR/TreeX4GD.desktop"
rm -f "$ICON_DIR/TreeX4GD.png"

echo "🔄 Refreshing system application caches..."

# Refresh caches so the app disappears immediately
if command -v update-desktop-database >/dev/null 2>&1; then update-desktop-database "$APP_DIR"; fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then gtk-update-icon-cache -f -t "$HOME/.local/share/icons/hicolor" >/dev/null 2>&1; fi
if command -v xdg-desktop-menu >/dev/null 2>&1; then xdg-desktop-menu forceupdate --mode user; fi
touch "$APP_DIR"

echo "✅ Uninstallation complete!"
echo "TreeX4GD has been successfully removed from your system."
echo ""
read -p "Press Enter to close..."
EOF

    # Make the generated scripts executable
    chmod +x "$FINAL_DIST_DIR/install.sh"
    chmod +x "$FINAL_DIST_DIR/uninstall.sh"
    
    echo "✅ Installer & Uninstaller files created!"
    echo "👉 You can run './install.sh' from the output folder to install it,"
    echo "   or simply run the 'TreeX4GD' binary directly without installing."
else
    echo "⏭️  Skipped generating installer files."
    echo "👉 You can run the 'TreeX4GD' binary directly without installing."
fi

echo ""
echo "🎉 BUILD PROCESS FINISHED!"
echo "📁 Output location: $FINAL_DIST_DIR"
echo "✅ Done!"
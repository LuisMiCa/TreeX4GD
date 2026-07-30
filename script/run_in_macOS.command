#!/bin/sh

# POSIX-compatible launcher for TreeX4GD.py (macOS)
echo "----------------------------------------"
echo "TreeX4GD — Python Script Loader for macOS"
echo "----------------------------------------"

# --------------------------------------------------
# 0. OS Validation
# --------------------------------------------------
OS_NAME=$(uname -s)
if [ "$OS_NAME" = "Linux" ]; then
    echo "❌ This script is for macOS. Please run the Linux version."
    read dummy
    exit 1
elif [ "$OS_NAME" != "Darwin" ]; then
    echo "⚠️  Warning: Unknown OS ($OS_NAME). Proceeding anyway..."
fi

# --------------------------------------------------
# 1. Setup Working Directory & Environment Variables
# --------------------------------------------------
cd "$(dirname "$0")" || {
  echo "❌ Failed to change directory."
  exit 1
}

VENV_DIR="venv"
VENV_PYTHON="$VENV_DIR/bin/python"
SCRIPT_FILE="TreeX4GD.py"
REQUIREMENTS="requirements.txt"
export PYTHONUNBUFFERED=1

# --------------------------------------------------
# 2. Pre-flight Checks
# --------------------------------------------------
# 2.0. Define a reusable failure function
fail_python_check() {
    echo "$1"
    echo "👉 Please install the official Python 3.11 or newer from:"
    echo "⬇️  https://www.python.org/downloads/macos/"
    echo "--------------------------------------------------"
    echo "Press Enter to exit..."
    read dummy
    exit 1
}

# 2.1. Check if Python 3 is installed at all
if ! command -v python3 >/dev/null 2>&1; then
    fail_python_check "❌ Python 3 is not installed."
fi

# 2.2. Check for macOS Stub (prevents the "Install Command Line Tools" popup)
# If the binary is the system stub AND Xcode tools are missing, fail gracefully.
if [ "$(command -v python3)" = "/usr/bin/python3" ] && ! xcode-select -p >/dev/null 2>&1; then
    fail_python_check "⚠️  System Python stub detected without Xcode Command Line Tools."
fi

# 2.3. Check Python version (must be 3.11+)
# We let Python handle the comparison logic. It returns exit code 1 if too old.
if ! python3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >/dev/null 2>&1; then
    CURRENT_VER=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
    fail_python_check "⚠️  Python version $CURRENT_VER is too old."
fi

# --------------------------------------------------
# 3. Virtual Environment Validation & Creation
# --------------------------------------------------
NEED_VENV=false

# 3.1. Check if venv folder exists
if [ -d "$VENV_DIR" ]; then
    # Check if the python executable actually exists inside it
    if [ -f "$VENV_PYTHON" ]; then
        # Check if it's executable and not corrupted
        if ! "$VENV_PYTHON" -c "import sys; sys.exit(0)" >/dev/null 2>&1; then
             echo "⚠️  Virtual environment appears corrupted."
             echo "♻️  Removing venv for clean rebuild..."
             rm -rf "$VENV_DIR"
             NEED_VENV=true
        else
             echo "✅ Virtual environment detected."
        fi
    else
        echo "⚠️  Venv folder exists but Python is missing."
        rm -rf "$VENV_DIR"
        NEED_VENV=true
    fi
else
    NEED_VENV=true
fi

# 3.2. Create Venv if missing or corrupted
if [ "$NEED_VENV" = "true" ]; then
  echo "📦 Creating virtual environment..."
  python3 -m venv "$VENV_DIR" || {
    echo "❌ Failed to create virtual environment."
    echo "Check out the error message(s) above for details."
    echo "Press Enter to exit..."
    read dummy
    exit 1
  }
  
  # Final sanity check
  if [ ! -f "$VENV_PYTHON" ]; then
      echo "❌ Venv creation reported success, but binary is missing."
      echo "Press Enter to exit..."
      read dummy
      exit 1
  fi
  echo "✅ Virtual environment created!"
fi

# --------------------------------------------------
# 4. Dependency Check & Installation
# --------------------------------------------------
echo "🔍 Checking dependencies..."

# 4.1. Verify that requirements.txt exists
if [ ! -f "$REQUIREMENTS" ]; then
    echo "❌ Error: '$REQUIREMENTS' not found!"
    echo "Press Enter to exit..."
    read dummy
    exit 1
fi

# 4.2. Check for missing dependencies using the venv's Python
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
MISSING_DEPS=$?

# 4.3. Check Internet Connectivity using the venv's Python
# Uses socket to verify an outward connection without triggering macOS SSL certificate errors.
# Returns 0 for success, non-zero for failure/timeout.
"$VENV_PYTHON" -c "import socket; socket.create_connection(('www.google.com', 443), timeout=3)" > /dev/null 2>&1
IS_ONLINE=$?

# 4.4. Decision Logic
if [ $MISSING_DEPS -ne 0 ]; then
    # Case: Dependencies are MISSING
    if [ $IS_ONLINE -ne 0 ]; then
        echo "❌ Dependencies are missing and you are OFFLINE."
        echo "   We cannot install the required libraries without internet."
        echo "   Please connect to the network and try again."
        echo "Press Enter to exit..."
        read dummy
        exit 1
    else
        echo "📦 Missing requirements detected."
        echo "⬇️  Installing dependencies..."
        "$VENV_PYTHON" -m pip install --upgrade pip --quiet
        "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS"
        if [ $? -ne 0 ]; then
            echo "❌ Installation failed."
            echo "Press Enter to exit..."
            read dummy
            exit 1
        fi
        echo "✅ Dependencies installed."
    fi
else
    # Case: Dependencies are PRESENT
    if [ $IS_ONLINE -ne 0 ]; then
        echo "⚠️  YOU ARE OFFLINE"
        echo "   All required libraries are installed, so the app will launch."
        echo "   However, features requiring internet access will not work,"
        echo "   but you can still process local JSON files."
        echo "--------------------------------------------------"
        # echo "Press Enter to launch anyway..."
        # read dummy
    else
        echo "✅ Dependencies satisfied & Internet connected."
    fi
fi

# --------------------------------------------------
# 5. Launch Application
# --------------------------------------------------
echo "🚀 Launching TreeX4GD..."
echo "--------------------------------------------------"

# Fix for rare Tkinter initialization issues on macOS
export TCL_LIBRARY="" 
export TK_LIBRARY=""

"$VENV_PYTHON" "$SCRIPT_FILE"
EXIT_CODE=$?

echo "--------------------------------------------------"
if [ $EXIT_CODE -ne 0 ]; then
    echo "⚠️  TreeX4GD closed with an error (Code: $EXIT_CODE)."
else
    echo "✅ TreeX4GD finished successfully."
fi

echo "Exiting in 2 seconds..."
sleep 2
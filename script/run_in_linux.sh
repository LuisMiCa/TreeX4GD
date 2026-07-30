#!/usr/bin/env bash

echo "----------------------------------------"
echo "TreeX4GD — Python Script Loader for Linux"
echo "----------------------------------------"

# --------------------------------------------------
# 0. OS Validation
# --------------------------------------------------
OS_NAME=$(uname -s)
if [ "$OS_NAME" = "Darwin" ]; then
    echo "❌ This script is for Linux. Please run the macOS version."
    read -p "Press Enter to exit..."
    exit 1
elif [ "$OS_NAME" != "Linux" ]; then
    echo "⚠️  Warning: Unknown OS ($OS_NAME). Proceeding anyway..."
fi

# --------------------------------------------------
# 1. Setup & Path Configuration
# --------------------------------------------------
cd "$(dirname "$0")" || {
    echo "❌ Failed to change directory."
    read -p "Press Enter to exit..."
    exit 1
}

# Define paths
VENV_DIR="venv"
VENV_PYTHON="$VENV_DIR/bin/python"
SCRIPT_FILE="TreeX4GD.py"
REQUIREMENTS="requirements.txt"
export PYTHONUNBUFFERED=1 

# --------------------------------------------------
# 2. Pre-flight Checks
# --------------------------------------------------
# 2.1. Check if main script exists
if [ ! -f "$SCRIPT_FILE" ]; then
    echo "❌ $SCRIPT_FILE not found in $(pwd)"
    read -p "Press Enter to exit..."
    exit 1
fi

# 2.2. Check if Python 3 is installed
if ! command -v python3 >/dev/null 2>&1; then
    echo "❌ python3 is not installed."
    echo "Please install it using your package manager."
    read -p "Press Enter to exit..."
    exit 1
fi

# 2.3. Version Check (Ensures Python 3.11+)
if ! python3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >/dev/null 2>&1; then
    CURRENT_VER=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
    echo "❌ Python version $CURRENT_VER is too old."
    echo "👉 This application requires Python 3.11 or newer."
    echo "   Please upgrade Python or install a newer version alongside it."
    read -p "Press Enter to exit..."
    exit 1
fi

# --------------------------------------------------
# 3. Check Tkinter (System Level)
# --------------------------------------------------
if ! python3 -c "import tkinter" >/dev/null 2>&1; then
    echo "❌ The 'tkinter' module is missing."
    if command -v apt >/dev/null; then
        echo "💡 Detected apt-based system (Ubuntu/Debian)."
        read -p "❓ Install 'python3-tk' now? (y/n): " choice
        if [[ "$choice" == [yY]* ]]; then
            echo "🔑 Enter sudo password if requested..."
            sudo apt update && sudo apt install -y python3-tk
        fi
    fi

    if ! python3 -c "import tkinter" >/dev/null 2>&1; then
        echo "⛔ Please install python3-tk manually."
        read -p "Press Enter to exit..."
        exit 1
    fi
    echo "✅ Tkinter is ready."
fi

# --------------------------------------------------
# 4. Virtual Environment Validation
# --------------------------------------------------
NEED_VENV=false

# 4.1. Check if venv directory and python executable exist
if [ -d "$VENV_DIR" ] && [ -f "$VENV_PYTHON" ]; then
    # Check if venv is not corrupted by running a simple Python command
    "$VENV_PYTHON" -c "import sys; sys.exit(0)" >/dev/null 2>&1
    if [ $? -ne 0 ]; then
        echo "⚠️  Virtual environment corrupted. Rebuilding..."
        rm -rf "$VENV_DIR"
        NEED_VENV=true
    else
        echo "✅ Virtual environment detected."
    fi
else
    NEED_VENV=true
fi

# 4.2. Create venv if needed
if [ "$NEED_VENV" = "true" ]; then
    echo "📦 Creating virtual environment..."
    python3 -m venv "$VENV_DIR"
    if [ $? -ne 0 ]; then
        echo "⚠️  Venv creation failed. Trying to fix missing venv package..."
        if command -v apt >/dev/null; then
             sudo apt update && sudo apt install -y python3-venv && python3 -m venv "$VENV_DIR"
        else
             echo "❌ Could not create venv. Please install python3-venv manually."
             read -p "Press Enter to exit..."
             exit 1
        fi
    fi
    
    if [ ! -f "$VENV_PYTHON" ]; then
        echo "❌ Failed to create virtual environment."
        read -p "Press Enter to exit..."
        exit 1
    fi
    echo "✅ Virtual environment created."
fi


# --------------------------------------------------
# 5. Dependency Check & Internet Validation
# --------------------------------------------------
echo "🔍 Checking dependencies..."
# 5.1. Verify requirements.txt exists
if [ ! -f "$REQUIREMENTS" ]; then
    echo "❌ Error: '$REQUIREMENTS' not found!"
    read -p "Press Enter to exit..."
    exit 1
fi

# 5.2. Check if dependencies are already installed
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

# 5.3. Check Internet Connectivity 
# Using socket to verify an outward connection without triggering SSL certificate errors.
# Returns 0 if successful, non-zero if it fails or times out.
"$VENV_PYTHON" -c "import socket; socket.create_connection(('www.google.com', 443), timeout=3)" > /dev/null 2>&1
IS_ONLINE=$?

# 5.4. Decision Logic
if [ $MISSING_DEPS -ne 0 ]; then
    # Case: Dependencies are MISSING
    if [ $IS_ONLINE -ne 0 ]; then
        echo "❌ Dependencies are missing, and you are OFFLINE."
        echo "   We cannot install the required libraries without internet."
        echo "   Please connect to the network and try again."
        read -p "Press Enter to exit..."
        exit 1
    else
        echo "📦 Missing requirements detected."
        echo "⬇️  Installing dependencies..."
        "$VENV_PYTHON" -m pip install --upgrade pip --quiet
        "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS"
        if [ $? -ne 0 ]; then
            echo "❌ Installation failed."
            read -p "Press Enter to exit..."
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
        # read -p "Press Enter to launch anyway..."
    else
        echo "✅ Dependencies satisfied & Internet connected."
    fi
fi


# --------------------------------------------------
# 6. Launch Application
# --------------------------------------------------
echo "🚀 Launching TreeX4GD..."
echo "--------------------------------------------------"

"$VENV_PYTHON" "$SCRIPT_FILE"
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo "--------------------------------------------------"
    echo "⚠️  App closed with error code: $EXIT_CODE"
    read -p "Press Enter to close..."
else
    echo "--------------------------------------------------"
    echo "✅ App finished successfully."
    sleep 2
fi
exit $EXIT_CODE
#!/bin/bash

# Excel2JSON for Mac - Run Script
# Usage: 
#   ./run.sh                    - Convert all Excel files
#   ./run.sh [files...]         - Convert specific files

echo "Excel2JSON for Mac"
echo "=================="

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "Error: Node.js is not installed. Please install Node.js first."
    echo "Visit: https://nodejs.org/"
    exit 1
fi

# Check if dependencies are installed
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install
fi

# Make the script executable
chmod +x excel2json-simple.js

# Determine mode and files
MODE="simple"
FILES=()

if [ $# -eq 0 ]; then
    # No arguments - convert all files in simple mode
    echo "Converting all Excel files in current directory (simple mode)..."
    node excel2json-simple.js
elif [ "$1" = "simple" ]; then
    # Simple mode with optional files
    shift
    if [ $# -eq 0 ]; then
        echo "Converting all Excel files in current directory (simple mode)..."
        node excel2json-simple.js
    else
        echo "Converting specified files (simple mode)..."
        node excel2json-simple.js "$@"
    fi
elif [ "$1" = "advanced" ]; then
    # Advanced mode not available - use simple mode instead
    echo "Advanced mode is not available. Using simple mode instead..."
    shift
    if [ $# -eq 0 ]; then
        echo "Converting all Excel files in current directory (simple mode)..."
        node excel2json-simple.js
    else
        echo "Converting specified files (simple mode)..."
        node excel2json-simple.js "$@"
    fi
else
    # Files specified without mode - use simple mode
    echo "Converting specified files (simple mode)..."
    node excel2json-simple.js "$@"
fi

echo "Done!"

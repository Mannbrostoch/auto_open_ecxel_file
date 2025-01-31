#!/bin/bash

# Check if input directory exists
if [ ! -d "input" ]; then
    echo "Error: input directory does not exist"
    exit 1
fi

# Create output directory if it doesn't exist
mkdir -p output

# Check for files with better error handling
if [ ! "$(ls -A input/*.xls 2>/dev/null)" ]; then
    echo "Error: No .xls files found in input directory"
    exit 1
fi

# Run AppleScript with error checking
echo "Processing Excel files..."
if ! osascript convert.scpt; then
    echo "Error: AppleScript execution failed"
    exit 1
fi

echo "Done processing files"
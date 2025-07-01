#!/bin/bash

# Initial setup script for DocMailer
# For Windows, run in Git Bash or WSL

echo "=== DocMailer Initial Setup ==="
echo

# Restore NuGet packages
echo "Restoring NuGet packages..."
dotnet restore

# Build the project
echo "Building the project..."
dotnet build

# Create directories if they don't exist
echo "Creating required directories..."
mkdir -p Data
mkdir -p Output
mkdir -p Templates

echo
echo "=== Setup Complete ==="
echo
echo "Next steps:"
echo "1. Edit config.json file with your email settings"
echo "2. Place your Excel file in Data/ folder"
echo "3. Run: dotnet run test (to test configuration)"
echo "4. Run: dotnet run process (to process recipients)"
echo

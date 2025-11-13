#!/bin/bash

# LibreOffice macOS setup script for non-technical users
# Auto-setup for PowerPoint Script Slide Converter

echo "=== PowerPoint Script Slide Converter Setup ==="
echo ""

# Download LibreOffice
echo "1. Downloading LibreOffice..."
cd ~/Downloads
curl -L -o LibreOffice.dmg "https://download.documentfoundation.org/libreoffice/stable/7.6.4/mac/x86_64/LibreOffice_7.6.4_MacOS_x86-64.dmg"

# Mount
echo "2. Mounting installer..."
hdiutil attach LibreOffice.dmg

# Copy to Applications
echo "3. Installing application..."
cp -r "/Volumes/LibreOffice/LibreOffice.app" /Applications/

# Unmount
echo "4. Cleaning up..."
hdiutil detach "/Volumes/LibreOffice"
rm LibreOffice.dmg

# Gatekeeper bypass (safe approach - no system changes)
echo "5. Adjusting security settings..."

# Remove quarantine from LibreOffice app only
sudo xattr -d com.apple.quarantine "/Applications/LibreOffice.app" 2>/dev/null || true
sudo xattr -d com.apple.quarantine "/Applications/LibreOffice.app/Contents/MacOS/soffice" 2>/dev/null || true
sudo xattr -d com.apple.quarantine "/Applications/LibreOffice.app/Contents/MacOS/LibreOffice" 2>/dev/null || true

echo "Note: First launch may show security prompt. Click 'Open' in the dialog."

# Install Python dependencies
echo "6. Installing Python dependencies..."
pip3 install -r requirements_web.txt

echo ""
echo "Setup complete!"
echo ""
echo "To start:"
echo "1. Open terminal and go to this folder: cd $(pwd)"
echo "2. Start server: python3 web_app.py"
echo "3. Open browser: http://localhost:8000"
echo ""
echo "Note: If you see a security warning on first launch,"
echo "go to System Preferences > Security & Privacy and allow it."

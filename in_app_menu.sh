#!/bin/bash

# Python app as app in APP MENU 

# Variables
USERNAME=$(whoami)
HOME_DIR="/home/$USERNAME"
SCRIPT_PATH="$HOME_DIR/Downloads/apps/video-frame-extractor/run_fxtr.sh"
DESKTOP_FILE="$HOME_DIR/.local/share/applications/fxtr.desktop"

# Step 1: Create the shell script
echo "Creating shell script to run fxtr..."
cat <<EOL > $SCRIPT_PATH
#!/bin/bash
source ~/py310env/bin/activate
python ~/Downloads/apps/video-frame-extractor/src/app.py
EOL
chmod +x $SCRIPT_PATH
echo "Shell script created at $SCRIPT_PATH."

# Step 2: Create the desktop entry
echo "Creating desktop entry for fxtr..."
mkdir -p ~/.local/share/applications
cat <<EOL > $DESKTOP_FILE
[Desktop Entry]
Version=1.0
Name=vfxtr
Comment=Run the Video Frame Extractor app
Exec=$SCRIPT_PATH
Icon=/usr/share/icons/hicolor/scalable/apps/fxtr.svg
Terminal=true
Type=Application
Categories=Utility;
EOL
echo "Desktop entry created at $DESKTOP_FILE."

# Step 3: Update desktop database
echo "Updating desktop database..."
update-desktop-database ~/.local/share/applications
echo "Desktop database updated."

# Final Message
echo "Setup complete! You can now find 'Video Frame Extractor' as 'vfxtr' in your application menu."

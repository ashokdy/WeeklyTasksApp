#!/bin/bash
# ─────────────────────────────────────────────────────────────────
#  MAG Weekly Tasks — Install as permanent macOS background service
#  Run once: bash install_service.sh
# ─────────────────────────────────────────────────────────────────

set -e

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
PLIST_NAME="com.mag.weeklytasks"
PLIST_DST="$HOME/Library/LaunchAgents/$PLIST_NAME.plist"
LOG_DIR="$APP_DIR/logs"

echo ""
echo "  MAG Weekly Tasks — Service Installer"
echo "  ──────────────────────────────────────"

# ── Find Python with Flask ─────────────────────────────────────────
PYTHON=""
for py in /usr/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
    if $py -c "import flask" 2>/dev/null; then
        PYTHON="$py"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo ""
    echo "  Flask not found. Installing dependencies first..."
    pip3 install -r "$APP_DIR/requirements.txt" -q
    PYTHON=$(which python3)
fi

echo "  Python  : $PYTHON"
echo "  App dir : $APP_DIR"
echo "  Plist   : $PLIST_DST"

# ── Create logs folder ─────────────────────────────────────────────
mkdir -p "$LOG_DIR"

# ── Write LaunchAgent plist ────────────────────────────────────────
mkdir -p "$HOME/Library/LaunchAgents"

cat > "$PLIST_DST" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.mag.weeklytasks</string>

    <key>ProgramArguments</key>
    <array>
        <string>$PYTHON</string>
        <string>$APP_DIR/app.py</string>
    </array>

    <key>WorkingDirectory</key>
    <string>$APP_DIR</string>

    <key>RunAtLoad</key>
    <true/>

    <key>KeepAlive</key>
    <true/>

    <key>StandardOutPath</key>
    <string>$LOG_DIR/out.log</string>

    <key>StandardErrorPath</key>
    <string>$LOG_DIR/err.log</string>

    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/usr/local/bin:/usr/bin:/bin</string>
    </dict>
</dict>
</plist>
PLIST

# ── Unload old instance if running ────────────────────────────────
launchctl unload "$PLIST_DST" 2>/dev/null || true

# ── Load and start ─────────────────────────────────────────────────
launchctl load -w "$PLIST_DST"

sleep 2

# ── Verify ────────────────────────────────────────────────────────
if launchctl list | grep -q "$PLIST_NAME"; then
    echo ""
    echo "  ✅  Service installed and running!"
    echo ""
    echo "  Open in browser → http://localhost:5000"
    echo ""
    echo "  Logs  : $LOG_DIR/out.log"
    echo "          $LOG_DIR/err.log"
    echo ""
    echo "  To stop the service:   bash uninstall_service.sh"
    echo ""
else
    echo ""
    echo "  ⚠️  Service loaded but may still be starting."
    echo "     Check logs: $LOG_DIR/err.log"
    echo ""
fi

open "http://localhost:5000" 2>/dev/null || true

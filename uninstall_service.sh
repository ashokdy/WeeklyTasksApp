#!/bin/bash
# ──────────────────────────────────────────────
#  MAG Weekly Tasks — Remove background service
# ──────────────────────────────────────────────

PLIST="$HOME/Library/LaunchAgents/com.mag.weeklytasks.plist"

if [ -f "$PLIST" ]; then
    launchctl unload -w "$PLIST" 2>/dev/null || true
    rm -f "$PLIST"
    echo "  ✅  Service stopped and removed."
else
    echo "  Service not found (already removed)."
fi

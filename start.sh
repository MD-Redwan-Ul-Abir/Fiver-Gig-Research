#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
#  Fiverr Research Hub — Container Startup
#  Boot order: Xvfb → Fluxbox → x11vnc → noVNC → Python GUI
# ─────────────────────────────────────────────────────────────────────────────

# Do NOT use set -e here — minor VNC/WM hiccups must not kill the container
set +e

# ── 1. Clean stale X lock files (left over from a forced restart) ─────────────
rm -f /tmp/.X1-lock /tmp/.X11-unix/X1 2>/dev/null

# ── 2. Virtual framebuffer ────────────────────────────────────────────────────
echo "  [1/5] Starting virtual display (Xvfb :1) …"
Xvfb :1 -screen 0 1280x900x24 -ac +extension GLX +render -noreset &
XVFB_PID=$!
sleep 2

# Verify Xvfb started
if ! kill -0 $XVFB_PID 2>/dev/null; then
    echo "  ERROR: Xvfb failed to start. Exiting."
    exit 1
fi

# ── 3. Window manager ─────────────────────────────────────────────────────────
echo "  [2/5] Starting window manager (Fluxbox) …"
DISPLAY=:1 fluxbox >/tmp/fluxbox.log 2>&1 &
sleep 1

# ── 4. VNC server ─────────────────────────────────────────────────────────────
echo "  [3/5] Starting VNC server (port 5900, no password) …"
x11vnc -display :1 -nopw -forever -shared -quiet -bg -o /tmp/x11vnc.log
sleep 1

# ── 5. noVNC web client ───────────────────────────────────────────────────────
echo "  [4/5] Starting noVNC web client (port 6080) …"

# Find the noVNC web-files directory (path differs by OS/package version)
NOVNC_DIR=""
for p in /usr/share/novnc /usr/share/noVNC /opt/novnc; do
    if [ -f "$p/vnc.html" ] || [ -f "$p/vnc_lite.html" ]; then
        NOVNC_DIR="$p"
        break
    fi
done

if [ -n "$NOVNC_DIR" ]; then
    websockify --web "$NOVNC_DIR" 6080 localhost:5900 >/tmp/novnc.log 2>&1 &
    sleep 1
    echo "  noVNC web UI : http://localhost:6080/vnc.html"
else
    # Fallback: raw websockify (VNC-over-WebSocket without web UI)
    websockify 6080 localhost:5900 >/tmp/novnc.log 2>&1 &
    echo "  noVNC web files not found — use a VNC client on port 5900 instead"
fi

# ── 6. Launch the GUI app ─────────────────────────────────────────────────────
echo "  [5/5] Launching Fiverr Research Hub …"
echo ""
echo "  ╔══════════════════════════════════════════════════════════╗"
echo "  ║          Fiverr Research Hub  —  READY                  ║"
echo "  ║                                                          ║"
echo "  ║  Browser access (recommended):                          ║"
echo "  ║    http://localhost:6080/vnc.html                       ║"
echo "  ║                                                          ║"
echo "  ║  Desktop VNC client (optional):                         ║"
echo "  ║    Host: localhost   Port: 5900   Password: (none)      ║"
echo "  ╚══════════════════════════════════════════════════════════╝"
echo ""

set -e
exec python /app/FiverrResearchHub.py

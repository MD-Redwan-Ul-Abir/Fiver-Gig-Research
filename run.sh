#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
#  Fiverr Research Hub — Launcher  (macOS / Linux)
#  Usage:
#    ./run.sh            — build (if needed) and start
#    ./run.sh --build    — force a full rebuild before starting
#    ./run.sh --stop     — stop the running container
#    ./run.sh --logs     — tail the container logs
# ─────────────────────────────────────────────────────────────────────────────
set -e

case "${1}" in
  --stop)
    echo "  Stopping Fiverr Research Hub…"
    docker-compose down
    exit 0
    ;;
  --logs)
    docker-compose logs -f
    exit 0
    ;;
  --build)
    echo "  Force-rebuilding image…"
    docker-compose build --no-cache
    ;;
esac

# ── Pre-flight: create required host-side files and folders ───────────────────
echo "  Checking required files and folders…"

mkdir -p "Excel and Images/gig_images"

# Docker mounts files, not directories, for these two —
# if they don't exist as FILES first, Docker creates DIRECTORIES instead
if [ ! -f api_config.json ]; then
    echo '{}' > api_config.json
    echo "  Created api_config.json  (add your API key inside the app)"
fi

if [ ! -f hub_config.json ]; then
    echo '{}' > hub_config.json
    echo "  Created hub_config.json"
fi

echo ""
echo "  Starting container…"
echo "  Once ready, open:  http://localhost:6080/vnc.html"
echo ""

docker-compose up

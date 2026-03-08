FROM python:3.11-slim

# ── Environment ───────────────────────────────────────────────────────────────
ENV DEBIAN_FRONTEND=noninteractive \
    DISPLAY=:1 \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

# ── Step 1: System packages for display, VNC, Tkinter, fonts ─────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    # Virtual framebuffer
    xvfb \
    # VNC server
    x11vnc \
    # Lightweight window manager
    fluxbox \
    # noVNC browser client + WebSocket proxy
    novnc \
    websockify \
    # Tkinter — must match the python:3.11 image's Python build
    python3-tk \
    tk-dev \
    # Fonts
    fonts-liberation \
    fonts-dejavu-core \
    fontconfig \
    # Utilities
    procps \
    curl \
    ca-certificates \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

# ── Step 2: Python dependencies ───────────────────────────────────────────────
WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

# ── Step 3: Playwright + Chromium + ALL its system dependencies ───────────────
# --with-deps automatically installs every OS library Chromium needs
RUN playwright install --with-deps chromium

# ── Step 4: Application code ──────────────────────────────────────────────────
COPY . .

# Pre-create the output directory tree inside the image
RUN mkdir -p "Excel and Images/gig_images"

# ── Step 5: Startup script ────────────────────────────────────────────────────
COPY start.sh /start.sh
RUN chmod +x /start.sh

# noVNC web UI  : http://localhost:6080/vnc.html
# Raw VNC       : localhost:5900  (optional desktop VNC client)
EXPOSE 6080 5900

CMD ["/start.sh"]

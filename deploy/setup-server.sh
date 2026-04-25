#!/bin/bash
# Evye Document Generator — Server Setup Script
# Run this ON the evye.co server after uploading the files.
#
# Usage:
#   1. SCP the deploy package to the server
#   2. SSH in and run: sudo bash /opt/evye-docgen/deploy/setup-server.sh

set -euo pipefail

INSTALL_DIR="/opt/evye-docgen"

echo "=== Evye Document Generator — Server Setup ==="

# 1. Create Python venv and install deps
echo "[1/4] Setting up Python virtual environment..."
cd "$INSTALL_DIR"
python3 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r deploy/requirements-server.txt

# 2. Create output directory
echo "[2/4] Creating output directory..."
mkdir -p "$INSTALL_DIR/output"
chown -R evyeadmin:evyeadmin "$INSTALL_DIR"

# 3. Install systemd service
echo "[3/4] Installing systemd service..."
cp deploy/evye-docgen.service /etc/systemd/system/evye-docgen.service
systemctl daemon-reload
systemctl enable evye-docgen
systemctl start evye-docgen

# 4. Add cleanup cron (delete files older than 7 days)
echo "[4/4] Adding cleanup cron..."
(crontab -l 2>/dev/null || true; echo "0 3 * * * find /opt/evye-docgen/output -name '*.docx' -mtime +7 -delete") | sort -u | crontab -

echo ""
echo "=== Setup complete ==="
echo ""
echo "Service status:"
systemctl status evye-docgen --no-pager || true
echo ""
echo "Test: curl http://127.0.0.1:8089/api/health"
echo ""
echo "NEXT STEP: Configure OpenLiteSpeed proxy in CyberPanel:"
echo "  1. Go to CyberPanel → Websites → evye.co → Rewrite Rules or vHost Conf"
echo "  2. Add a proxy context:"
echo "     URI: /docs/"
echo "     Type: Proxy"
echo "     Web Server: http://127.0.0.1:8089/"
echo "  3. Restart OpenLiteSpeed"
echo ""
echo "Then test: https://evye.co/docs/"

#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────
#  setup_cron.sh
#  Installs a daily cron job for the Process Pipeline.
#  Run once: bash setup_cron.sh
# ─────────────────────────────────────────────────────────────────

set -euo pipefail

# ── Config ────────────────────────────────────────────────────────
# Edit RUN_HOUR/RUN_MIN to change the daily run time (server local time).
RUN_HOUR=7        # 07:00 AM  — change as needed
RUN_MIN=0

# Resolve absolute path of this script's directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
NODE_BIN="$(command -v node)"
LOG_FILE="${SCRIPT_DIR}/logs/process.log"
# ─────────────────────────────────────────────────────────────────

# Create logs directory if missing
mkdir -p "${SCRIPT_DIR}/logs"

# Cron entry — runs every day at RUN_HOUR:RUN_MIN
CRON_CMD="${RUN_MIN} ${RUN_HOUR} * * * cd \"${SCRIPT_DIR}\" && \"${NODE_BIN}\" process.js >> \"${LOG_FILE}\" 2>&1"

echo ""
echo "──────────────────────────────────────────────────────────"
echo "  Process Pipeline — Cron Job Installer"
echo "──────────────────────────────────────────────────────────"
echo "  Script dir : ${SCRIPT_DIR}"
echo "  Node binary: ${NODE_BIN}"
echo "  Log file   : ${LOG_FILE}"
echo "  Schedule   : daily at ${RUN_HOUR}:$(printf '%02d' ${RUN_MIN}) (server local time)"
echo "──────────────────────────────────────────────────────────"
echo ""

# Check if entry already exists
if crontab -l 2>/dev/null | grep -qF "process.js"; then
  echo "⚠  A cron job for process.js already exists:"
  crontab -l 2>/dev/null | grep "process.js"
  echo ""
  read -rp "Replace it? [y/N] " confirm
  if [[ "${confirm,,}" != "y" ]]; then
    echo "Aborted — no changes made."
    exit 0
  fi
  # Remove existing entry
  crontab -l 2>/dev/null | grep -vF "process.js" | crontab -
fi

# Append new cron entry
( crontab -l 2>/dev/null; echo "${CRON_CMD}" ) | crontab -

echo "✓  Cron job installed:"
echo "   ${CRON_CMD}"
echo ""
echo "   Verify with:  crontab -l"
echo "   View logs  :  tail -f ${LOG_FILE}"
echo ""
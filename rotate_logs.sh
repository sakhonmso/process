#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────
#  rotate_logs.sh
#  Keeps only the last 30 days of log lines in p4p_merge.log.
#  Add to crontab (monthly) if the log grows large:
#    0 0 1 * * bash /path/to/p4p-excel-merge/rotate_logs.sh
# ─────────────────────────────────────────────────────────────────

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_FILE="${SCRIPT_DIR}/logs/p4p_merge.log"

if [[ ! -f "${LOG_FILE}" ]]; then
  echo "No log file found at ${LOG_FILE}"
  exit 0
fi

LINES_BEFORE=$(wc -l < "${LOG_FILE}")
# Keep last 5000 lines (~30 days of daily runs with verbose output)
tail -n 5000 "${LOG_FILE}" > "${LOG_FILE}.tmp" && mv "${LOG_FILE}.tmp" "${LOG_FILE}"
LINES_AFTER=$(wc -l < "${LOG_FILE}")

echo "Log rotated: ${LINES_BEFORE} → ${LINES_AFTER} lines (${LOG_FILE})"

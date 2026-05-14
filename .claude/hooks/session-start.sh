#!/bin/bash
# Session start hook: install Python dependencies for legal-document
# generation scripts (python-docx).
#
# - Synchronous mode: blocks session start until dependencies are ready,
#   guaranteeing scripts/build_*.py work immediately.
# - Web-only: skips when not running in Claude Code on the web.
# - Idempotent: pip install is safe to re-run.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-$(pwd)}"

python3 -m pip install --quiet --no-cache-dir -r requirements.txt

python3 -c "import docx" >/dev/null

echo "session-start: python-docx ready"

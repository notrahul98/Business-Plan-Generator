#!/usr/bin/env bash
# Runs every time the codespace starts: sets the sign-in password and starts the app on port 8400.
# Plans are stored in ./data inside this codespace only (never committed: see .gitignore).
set -euo pipefail
cd "$(dirname "$0")/.."
DATA="${KNS_DATA_DIR:-$PWD/data}"
mkdir -p "$DATA"

if [ -n "${KNS_APP_PASSWORD:-}" ]; then
  python -c "import os; from kns_plan.web.auth import hash_password; print(hash_password(os.environ['KNS_APP_PASSWORD']))" \
    > "$DATA/password.hash"
  rm -f "$DATA/FIRST-PASSWORD.txt"
  echo "Sign-in password: the KNS_APP_PASSWORD Codespaces secret."
elif [ ! -s "$DATA/password.hash" ]; then
  python - "$DATA" <<'PY'
import secrets, sys
from pathlib import Path
from kns_plan.web.auth import hash_password
data = Path(sys.argv[1])
pw = secrets.token_urlsafe(12)
(data / "password.hash").write_text(hash_password(pw))
(data / "FIRST-PASSWORD.txt").write_text(pw + "\n")
PY
  echo "Sign-in password generated: see data/FIRST-PASSWORD.txt"
fi

pkill -f "kns_plan.cli serve" 2>/dev/null || true
KNS_DATA_DIR="$DATA" KNS_HTTPS=1 setsid nohup python -m kns_plan.cli serve --host 127.0.0.1 --port 8400 \
  > "$DATA/server.log" 2>&1 < /dev/null &
echo "Business Plan Generator is starting on port 8400 (log: data/server.log). Open the Ports tab if no browser tab appears."

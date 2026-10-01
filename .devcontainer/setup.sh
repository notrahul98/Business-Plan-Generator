#!/usr/bin/env bash
# Runs once when the codespace is created: LibreOffice (for the pre-download check) and the app.
set -euo pipefail
cd "$(dirname "$0")/.."

sudo apt-get update -qq
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y -qq --no-install-recommends libreoffice-calc-nogui
python -m pip install -q --upgrade pip
python -m pip install -q -e ".[web,dev]"
echo "Setup finished: $(soffice --version 2>/dev/null | head -1)"

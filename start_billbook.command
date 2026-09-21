#!/bin/zsh
# Double-click this file on macOS, or run it from Terminal, to start the local app.
set -e
cd "$(dirname "$0")"
source .venv/bin/activate
export PYTHONPATH="src${PYTHONPATH:+:$PYTHONPATH}"
python -m nv_billbook.webapp

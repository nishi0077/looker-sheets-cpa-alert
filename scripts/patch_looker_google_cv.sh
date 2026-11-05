#!/bin/zsh
set -euo pipefail
export PATH="/Users/nishi/.local/bin:/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin"
cd "/Users/nishi/Cursol/Gravity Craft"
/Users/nishi/.local/bin/uv run --no-project --with google-auth --with google-auth-httplib2 --with google-api-python-client python scripts/patch_looker_google_cv.py

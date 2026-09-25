#!/usr/bin/env bash
# Preview the app on this PC before pushing.
#   ./preview.sh          → http://localhost:8080        (sign in, your real data)
#                           http://localhost:8080/?demo  (sample data, no sign-in)
#   ./preview.sh 9000     → same, on another port
set -e
cd "$(dirname "$0")"
PORT="${1:-8080}"

# The Firebase web config isn't secret (the live site serves it), so borrow it
# from the deployed site the first time. config.js is gitignored.
if [ ! -s config.js ]; then
  if curl -fsS https://ijishnupr.github.io/jeournal/config.js -o config.js; then
    echo "Fetched config.js from the live site."
  else
    rm -f config.js
    echo "No config.js — sign-in won't work, but ?demo will. (Copy config.example.js to config.js to fix.)"
  fi
fi

echo
echo "  Your data:   http://localhost:$PORT/"
echo "  Sample data: http://localhost:$PORT/?demo#stats"
echo
echo "Edit index.html and just reload the browser. Ctrl+C to stop."
exec python3 -m http.server "$PORT" --bind 127.0.0.1

#!/bin/zsh
# Auto-push: commit any changes in this repo and push to origin/main.
# Invoked every 10 minutes by a launchd job. No-op when there are no changes.

REPO="/Users/maxvulcanx/Desktop/AI Project/COE AI Creative Economy"
cd "$REPO" || exit 1

# Ensure git can find a sane PATH when run by launchd
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"

# Nothing to do if working tree is clean
if [ -z "$(git status --porcelain)" ]; then
  exit 0
fi

git add -A
git commit -q -m "Auto-push: $(date '+%Y-%m-%d %H:%M:%S')" || exit 0
git push -q origin main

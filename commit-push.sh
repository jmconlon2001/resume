#!/bin/bash
set -e

cd "$(dirname "$0")"

git add -A
git status

read -p "Commit message: " msg
if [ -z "$msg" ]; then
  echo "Commit message cannot be empty. Aborting."
  exit 1
fi

git commit -m "$msg"
git push

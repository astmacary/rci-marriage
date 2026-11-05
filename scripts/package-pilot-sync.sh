#!/bin/sh
set -eu
if [ "$#" -ne 1 ]; then
  echo "Usage: sh scripts/package-pilot-sync.sh /absolute/path/to/pilot-sync.zip" >&2
  exit 2
fi
case "$1" in
  /*.zip) output=$1 ;;
  *) echo "Provide an absolute .zip path" >&2; exit 2 ;;
esac
if [ -e "$output" ]; then
  echo "Choose a new output path" >&2
  exit 2
fi
cd "$(dirname "$0")/.."
stage=$(mktemp -d)
trap 'rm -r "$stage"' EXIT
mkdir "$stage/scripts" "$stage/src"
cp -R deploy/pilot-sync/. "$stage/"
cp scripts/pilot-sync.cjs "$stage/scripts/"
cp src/data.js "$stage/src/"
cd "$stage"
zip -q -r "$output" host.json PilotSync scripts src

#!/usr/bin/env bash
#
# Keep the README download link pinned to the current add-in version.
#
# Reads ADDIN_VERSION from src/Core/modStartup.bas and, if it differs,
# rewrites the version on the marked line in README.md (the display text
# and the release-asset URL). Idempotent and line-ending preserving - a
# no-op when they already match, so it is safe to run by hand or from the
# PostToolUse hook in .claude/settings.json that fires after
# src/Core/modStartup.bas is edited.
#
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
startup="$root/src/Core/modStartup.bas"
readme="$root/README.md"
marker='<!-- download-link: version managed by tools/sync-readme-version.sh -->'

[ -f "$startup" ] && [ -f "$readme" ] || exit 0

ver="$(grep -oE 'ADDIN_VERSION[^"]*"[0-9]+\.[0-9]+\.[0-9]+"' "$startup" \
       | grep -oE '[0-9]+\.[0-9]+\.[0-9]+' | head -1)"
[ -n "$ver" ] || { echo "sync-readme-version: no ADDIN_VERSION found" >&2; exit 0; }

grep -qF "$marker" "$readme" || { echo "sync-readme-version: marker not in README" >&2; exit 0; }

# Version currently on the line right after the marker.
current="$(grep -A1 -F "$marker" "$readme" | tail -1 \
           | grep -oE 'v[0-9]+\.[0-9]+\.[0-9]+' | head -1 || true)"
[ "$current" = "v$ver" ] && exit 0

# Rewrite just that one line; sed keeps each line's existing CRLF/LF.
sed -i "\#${marker}#{n;s/v[0-9]\{1,\}\.[0-9]\{1,\}\.[0-9]\{1,\}/v${ver}/g}" "$readme"
echo "sync-readme-version: README download link ${current:-<none>} -> v$ver"

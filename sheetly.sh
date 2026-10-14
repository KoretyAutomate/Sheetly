#!/usr/bin/env bash
cd "$(dirname "$(readlink -f "$0")")" && [ -d dist ] || npx vite build
exec npx electron . --no-sandbox "$@"

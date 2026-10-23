#!/bin/zsh
set -euo pipefail

exec node "${0:A:h}/server.js"

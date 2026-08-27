#!/bin/sh
set -eu

mkdir -p "${NUMBA_CACHE_DIR:-/tmp/numba}" "${XDG_CACHE_HOME:-/tmp/cache}"

echo "Preparing background-removal model: ${REMBG_MODEL:-u2netp}"
if python -m services.background_remover --prepare; then
    echo "Background-removal model is ready"
else
    echo "WARNING: background-removal model preparation failed; starting the bot without it" >&2
    echo "Other services will remain available. Check the container logs before using background removal." >&2
fi
echo "Starting Telegram services bot"
exec "$@"

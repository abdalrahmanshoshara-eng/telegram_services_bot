from pathlib import Path


def test_docker_writable_cache_configuration():
    dockerfile = Path("Dockerfile").read_text(encoding="utf-8")
    compose = Path("docker-compose.yml").read_text(encoding="utf-8")
    entrypoint = Path("docker-entrypoint.sh").read_text(encoding="utf-8")

    assert "NUMBA_CACHE_DIR=/tmp/numba" in dockerfile
    assert "U2NET_HOME=/data/rembg" in dockerfile
    assert "NUMBA_CACHE_DIR: /tmp/numba" in compose
    assert "U2NET_HOME: /data/rembg" in compose
    assert "XDG_CACHE_HOME: /tmp/cache" in compose
    assert 'mkdir -p "${NUMBA_CACHE_DIR:-/tmp/numba}"' in entrypoint
    assert "if python -m services.background_remover --prepare" in entrypoint

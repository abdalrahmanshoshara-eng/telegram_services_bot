import argparse
import os
import threading
from pathlib import Path

from PIL import Image

DEFAULT_MODEL = "u2netp"
_SESSION = None
_SESSION_MODEL = None
_SESSION_LOCK = threading.Lock()


def configured_model() -> str:
    return os.environ.get("REMBG_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL


def get_session():
    global _SESSION, _SESSION_MODEL
    model_name = configured_model()
    with _SESSION_LOCK:
        if _SESSION is None or _SESSION_MODEL != model_name:
            from rembg import new_session

            _SESSION = new_session(model_name)
            _SESSION_MODEL = model_name
    return _SESSION


def prepare_background_model() -> None:
    """Download (when missing) and load the configured model."""
    get_session()


def remove_background(source: Path, destination: Path) -> None:
    from rembg import remove

    session = get_session()
    with Image.open(source) as image:
        image.load()
        result = remove(image.convert("RGBA"), session=session)
        result.save(destination, format="PNG", optimize=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare rembg or process one image.")
    parser.add_argument("source", nargs="?", type=Path)
    parser.add_argument("destination", nargs="?", type=Path)
    parser.add_argument(
        "--prepare",
        action="store_true",
        help="Download and validate the configured background-removal model.",
    )
    args = parser.parse_args()

    if args.prepare:
        prepare_background_model()
        print(f"Background-removal model is ready: {configured_model()}")
        return
    if args.source is None or args.destination is None:
        parser.error("source and destination are required unless --prepare is used")
    remove_background(args.source, args.destination)


if __name__ == "__main__":
    main()

import sys
from types import SimpleNamespace

from PIL import Image

import services.background_remover as background_remover


def test_saves_png_with_alpha_channel(tmp_path, monkeypatch):
    source = tmp_path / "source.jpg"
    destination = tmp_path / "result.png"
    Image.new("RGB", (4, 4), "white").save(source)
    fake_rembg = SimpleNamespace(
        new_session=lambda model_name: f"session:{model_name}",
        remove=lambda image, session: image,
    )
    monkeypatch.setitem(sys.modules, "rembg", fake_rembg)
    monkeypatch.setattr(background_remover, "_SESSION", None)
    monkeypatch.setattr(background_remover, "_SESSION_MODEL", None)

    background_remover.remove_background(source, destination)

    with Image.open(destination) as result:
        assert result.format == "PNG"
        assert result.mode == "RGBA"

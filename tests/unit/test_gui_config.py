import json
from pathlib import Path

from src.core.mrp_gui import GUIConfig


def test_gui_config_persists_and_loads(tmp_path, monkeypatch):
    monkeypatch.setattr("src.core.mrp_gui.Path.home", lambda: tmp_path)

    config = GUIConfig(theme="darkly", window_size=(1024, 700), min_window_size=(800, 500))
    config.save()

    saved_payload = json.loads((tmp_path / ".mrp_analyzer" / "config.json").read_text())
    assert saved_payload["config_dir"] == str(tmp_path / ".mrp_analyzer")
    assert saved_payload["window_size"] == [1024, 700]

    loaded = GUIConfig.load()
    assert loaded.theme == "darkly"
    assert loaded.window_size == (1024, 700)
    assert loaded.min_window_size == (800, 500)


def test_gui_config_load_falls_back_to_default_on_invalid_json(tmp_path, monkeypatch):
    monkeypatch.setattr("src.core.mrp_gui.Path.home", lambda: tmp_path)
    config_dir = tmp_path / ".mrp_analyzer"
    config_dir.mkdir(parents=True)
    (config_dir / "config.json").write_text("{invalid json")

    loaded = GUIConfig.load()
    assert isinstance(loaded, GUIConfig)
    assert loaded.theme == "flatly"

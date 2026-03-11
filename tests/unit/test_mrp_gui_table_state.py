from pathlib import Path

import pandas as pd

from src.core.mrp_gui import AppState, GUIConfig, MRPGUI


class DummyVar:
    def __init__(self, value: str = ""):
        self._value = value

    def get(self) -> str:
        return self._value



def _build_gui_stub(page_size: int = 2) -> MRPGUI:
    gui = MRPGUI.__new__(MRPGUI)
    gui.state = AppState(config=GUIConfig(page_size=page_size))
    gui.selected_file = DummyVar("/tmp/input.xlsx")
    gui.column_box = {}
    gui._log = lambda *args, **kwargs: None
    return gui


def test_load_table_updates_state_and_pagination(monkeypatch):
    gui = _build_gui_stub(page_size=2)
    df = pd.DataFrame(
        {
            "CÓD": ["1", "2", "3", "4", "5"],
            "QUANTIDADE A SOLICITAR": [10, 20, 30, 40, 50],
            "ESTOQUE DISPONÍVEL": [1, 2, 3, 4, 5],
        }
    )
    render_calls = []

    monkeypatch.setattr("src.core.mrp_gui.pd.read_excel", lambda *args, **kwargs: df)
    gui._render_table = lambda: render_calls.append(gui.state.current_page)

    gui._load_table(path=Path("/tmp/itens_criticos.xlsx"))

    assert gui.state.df_table.equals(df)
    assert gui.state.current_page == 0
    assert gui.state.total_pages == 3
    assert gui.column_box["values"] == list(df.columns)
    assert render_calls == [0]

    gui._next_page()
    gui._next_page()
    gui._next_page()
    assert gui.state.current_page == 2

    gui._prev_page()
    gui._prev_page()
    gui._prev_page()
    assert gui.state.current_page == 0


def test_exports_use_state_dataframe(monkeypatch, tmp_path):
    gui = _build_gui_stub()
    gui.state.df_table = pd.DataFrame({"CÓD": ["A", "B"], "QUANTIDADE A SOLICITAR": [1, 2]})

    csv_target = tmp_path / "critical.csv"
    monkeypatch.setattr("src.core.mrp_gui.filedialog.asksaveasfilename", lambda **kwargs: str(csv_target))
    monkeypatch.setattr("src.core.mrp_gui.messagebox.showinfo", lambda *args, **kwargs: None)

    gui._export_csv()
    saved_csv = pd.read_csv(csv_target)
    assert saved_csv["CÓD"].tolist() == ["A", "B"]

    excel_target = tmp_path / "critical.xlsx"
    captured = {}

    def fake_to_excel(self, path, index=False):
        captured["self_id"] = id(self)
        captured["path"] = path
        captured["index"] = index

    monkeypatch.setattr("src.core.mrp_gui.filedialog.asksaveasfilename", lambda **kwargs: str(excel_target))
    monkeypatch.setattr(pd.DataFrame, "to_excel", fake_to_excel)

    gui._export_excel()

    assert captured["self_id"] == id(gui.state.df_table)
    assert captured["path"] == str(excel_target)
    assert captured["index"] is False

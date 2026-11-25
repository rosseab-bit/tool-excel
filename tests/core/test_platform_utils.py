import subprocess
from pathlib import Path

import pytest

from excel_tool import platform_utils as pu


@pytest.mark.parametrize(
    ("entrada", "esperado"),
    [
        ('a<b>c:d"e/f\\g|h?i*j.csv', "a_b_c_d_e_f_g_h_i_j.csv"),
        ("con\x00trol\x1f.txt", "con_trol_.txt"),
        ("CON", "CON_"),
        ("con.txt", "con_.txt"),
        ("Lpt9.csv", "Lpt9_.csv"),
        ("COM1", "COM1_"),
        ("COM10.txt", "COM10.txt"),
        ("console.txt", "console.txt"),
        ("", "sin_nombre"),
        ("...", "sin_nombre"),
        ("   ", "sin_nombre"),
        ("informe. . ", "informe"),
        ("Ñandú acción.xlsx", "Ñandú acción.xlsx"),
    ],
)
def test_safe_filename(entrada, esperado):
    assert pu.safe_filename(entrada) == esperado


def test_safe_filename_replacement_propio():
    assert pu.safe_filename("a/b", "-") == "a-b"
    assert pu.safe_filename("NUL.txt", "-") == "NUL-.txt"


def test_safe_filename_largo_conserva_extension():
    r = pu.safe_filename("a" * 400 + ".xlsx")
    assert len(r) == 150
    assert r.endswith(".xlsx")


def test_safe_filename_largo_sin_extension():
    assert len(pu.safe_filename("b" * 400)) == 150


def test_user_data_dir_linux_xdg(monkeypatch, tmp_path):
    monkeypatch.setattr(pu.sys, "platform", "linux")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "xdg"))
    d = pu.user_data_dir()
    assert d == tmp_path / "xdg" / "ExcelJoinTool"
    assert d.is_dir()
    assert pu.log_dir() == d / "logs"
    assert (d / "logs").is_dir()


def test_user_data_dir_linux_home(monkeypatch, tmp_path):
    monkeypatch.setattr(pu.sys, "platform", "linux")
    monkeypatch.delenv("XDG_DATA_HOME", raising=False)
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    assert pu.user_data_dir() == tmp_path / ".local" / "share" / "ExcelJoinTool"


def test_user_data_dir_windows_localappdata(monkeypatch, tmp_path):
    monkeypatch.setattr(pu.sys, "platform", "win32")
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    assert pu.user_data_dir() == tmp_path / "local" / "ExcelJoinTool"


def test_user_data_dir_windows_sin_variable(monkeypatch, tmp_path):
    monkeypatch.setattr(pu.sys, "platform", "win32")
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setattr(pu.Path, "home", classmethod(lambda cls: tmp_path))
    assert pu.user_data_dir() == tmp_path / "AppData" / "Local" / "ExcelJoinTool"


def test_open_folder_linux(monkeypatch, tmp_path):
    llamadas = []
    monkeypatch.setattr(pu.sys, "platform", "linux")
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **k: llamadas.append((a, k)))
    pu.open_folder(tmp_path)
    assert llamadas[0][0][0] == ["xdg-open", str(tmp_path)]
    assert "shell" not in llamadas[0][1]


def test_open_folder_linux_falla(monkeypatch, tmp_path):
    def falla(*a, **k):
        raise FileNotFoundError("xdg-open")

    monkeypatch.setattr(pu.sys, "platform", "linux")
    monkeypatch.setattr(subprocess, "Popen", falla)
    with pytest.raises(OSError, match="No se pudo abrir la carpeta"):
        pu.open_folder(tmp_path)


def test_open_folder_windows(monkeypatch, tmp_path):
    llamadas = []
    monkeypatch.setattr(pu.sys, "platform", "win32")
    monkeypatch.setattr(pu.os, "startfile", llamadas.append, raising=False)
    pu.open_folder(Path(tmp_path))
    assert llamadas == [tmp_path]

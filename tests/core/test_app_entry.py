import os
import sys

import pytest

from excel_tool import APP_NAME, __version__, app_entry


@pytest.fixture(autouse=True)
def _aislar(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "platform", sys.platform)
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "xdg"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    for h in list(app_entry.logger.handlers):
        h.close()
        app_entry.logger.removeHandler(h)
    yield
    for h in list(app_entry.logger.handlers):
        h.close()
        app_entry.logger.removeHandler(h)


def test_version(capsys):
    assert app_entry.main(["--version"]) == 0
    assert capsys.readouterr().out.strip() == f"{APP_NAME} {__version__}"


def test_version_sin_stdout(monkeypatch):
    monkeypatch.setattr(sys, "stdout", None)
    assert app_entry.main(["--version"]) == 0


def test_argumento_desconocido(capsys):
    assert app_entry.main(["--nada"]) == 2
    assert "Uso:" in capsys.readouterr().err


def test_argumento_desconocido_sin_stderr(monkeypatch):
    monkeypatch.setattr(sys, "stderr", None)
    assert app_entry.main(["--nada"]) == 2


def test_log_a_archivo(tmp_path):
    app_entry.main(["--version"])
    for h in app_entry.logger.handlers:
        h.flush()
    logs = list((tmp_path / "xdg" / "ExcelJoinTool" / "logs").glob("excel_tool.log"))
    assert logs
    assert __version__ in logs[0].read_text(encoding="utf-8")


def test_sin_carpeta_de_logs_no_falla(monkeypatch):
    def falla():
        raise PermissionError("no")

    monkeypatch.setattr(app_entry.platform_utils, "log_dir", falla)
    assert app_entry.main(["--version"]) == 0


def test_pasos_sin_pantalla():
    assert app_entry._paso_limites()
    assert app_entry._paso_archivos()


def test_self_test_marca_falla_si_no_hay_ui(monkeypatch, capsys):
    monkeypatch.setattr(app_entry, "_paso_ventana", lambda: "ok")
    monkeypatch.setitem(sys.modules, "excel_tool.ui.app", None)  # fuerza ImportError
    assert app_entry.main(["--self-test"]) == 1
    assert "FALLA" in capsys.readouterr().out


@pytest.mark.pantalla
def test_paso_ventana():
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("sin pantalla")
    assert app_entry._paso_ventana() == "ventana creada"

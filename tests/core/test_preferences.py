from pathlib import Path

import pytest

from excel_tool import platform_utils
from excel_tool.core.io import CsvOptions
from excel_tool.core.preferences import Preferences, load_preferences, save_preferences


@pytest.fixture
def datos(tmp_path, monkeypatch):
    carpeta = tmp_path / "datos"
    monkeypatch.setattr(platform_utils, "user_data_dir", lambda: carpeta)
    return carpeta


def test_ida_y_vuelta(datos, tmp_path):
    prefs = Preferences(
        last_folder=tmp_path / "mis archivos",
        last_csv_options=CsvOptions(separator=";", encoding="cp1252"),
        last_destination="excel",
    )
    save_preferences(prefs)  # la carpeta de datos no existía
    assert load_preferences() == prefs
    assert [p.name for p in datos.iterdir()] == ["preferencias.json"]


def test_archivo_inexistente(datos):
    assert load_preferences() == Preferences()


def test_carpeta_inexistente(tmp_path, monkeypatch):
    monkeypatch.setattr(platform_utils, "user_data_dir", lambda: tmp_path / "a" / "b")
    assert load_preferences() == Preferences()


@pytest.mark.parametrize(
    "contenido", ["{no es json", "[1, 2]", '{"last_csv_options": {"separator": ";"}}']
)
def test_archivo_danado(datos, contenido, caplog):
    datos.mkdir()
    (datos / "preferencias.json").write_text(contenido, encoding="utf-8")
    with caplog.at_level("WARNING"):
        assert load_preferences() == Preferences()
    assert caplog.records


def test_valores_nulos_y_unicode(datos):
    prefs = Preferences(last_folder=Path("ñandú"))
    save_preferences(prefs)
    assert load_preferences() == prefs


def test_guardar_no_rompe_si_falla(datos, monkeypatch, caplog):
    def boom():
        raise OSError("sin permisos")

    monkeypatch.setattr(platform_utils, "user_data_dir", boom)
    save_preferences(Preferences())
    assert caplog.records

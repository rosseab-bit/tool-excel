import sys
import types

import pytest

from excel_tool import platform_utils as pu
from excel_tool.core import memory
from excel_tool.core.memory import BASE_BYTES, check_memory

MB = 1024**2
GB = 1024**3


def _archivo(tmp_path, nombre, tam):
    p = tmp_path / nombre
    with p.open("wb") as f:
        f.truncate(tam)
    return p


def test_available_memory_linux_real():
    if not sys.platform.startswith("linux"):
        pytest.skip("solo Linux")
    valor = pu.available_memory()
    assert isinstance(valor, int)
    assert valor > 0


def test_available_memory_windows_simulado(monkeypatch):
    class Kernel:
        def GlobalMemoryStatusEx(self, ref):  # noqa: N802
            ref._obj.ullAvailPhys = 123456789
            return 1

    falso = types.SimpleNamespace(kernel32=Kernel())
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(pu.ctypes, "windll", falso, raising=False)
    assert pu.available_memory() == 123456789


def test_available_memory_windows_falla(monkeypatch):
    class Kernel:
        def GlobalMemoryStatusEx(self, ref):  # noqa: N802
            return 0

    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(
        pu.ctypes, "windll", types.SimpleNamespace(kernel32=Kernel()), raising=False
    )
    assert pu.available_memory() is None


def test_available_memory_otro_sistema(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    assert pu.available_memory() is None


def test_available_memory_linux_sin_meminfo(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setattr(pu, "_available_memory_linux", lambda: (_ for _ in ()).throw(OSError("x")))
    assert pu.available_memory() is None


@pytest.mark.parametrize(
    ("nombre", "factor"),
    [
        ("a.csv", 6),
        ("a.txt", 6),
        ("a.xlsx", 25),
        ("a.xlsm", 25),
        ("a.xlsb", 25),
        ("a.ods", 25),
        ("a.xls", 20),
        ("A.CSV", 6),
    ],
)
def test_factor_por_formato(tmp_path, monkeypatch, nombre, factor):
    monkeypatch.setattr(pu, "available_memory", lambda: 100 * GB)
    r = check_memory(_archivo(tmp_path, nombre, 10 * MB))
    assert r.needed == BASE_BYTES + 10 * MB * factor
    assert r.ok


def test_formato_explicito(tmp_path, monkeypatch):
    monkeypatch.setattr(pu, "available_memory", lambda: 100 * GB)
    r = check_memory(_archivo(tmp_path, "datos.dat", MB), format="xls")
    assert r.needed == BASE_BYTES + MB * 20


def test_alcanza_y_no_alcanza(tmp_path, monkeypatch):
    p = _archivo(tmp_path, "g.csv", 300 * MB)  # necesita 114 MB + 1800 MB
    monkeypatch.setattr(pu, "available_memory", lambda: 3 * GB)
    assert check_memory(p).ok
    monkeypatch.setattr(pu, "available_memory", lambda: 2 * GB)  # 0,8 * 2 GB = 1,6 GB
    r = check_memory(p)
    assert not r.ok
    assert r.available == 2 * GB
    assert "2,0 GB libres" in r.message
    assert "1,9 GB" in r.message
    assert r.message.startswith("Este archivo es muy grande")
    assert r.message.endswith("Cerrá otros programas y probá de nuevo.")


def test_margen(tmp_path, monkeypatch):
    p = _archivo(tmp_path, "g.csv", 100 * MB)  # 114 + 600 = 714 MB
    monkeypatch.setattr(pu, "available_memory", lambda: 800 * MB)
    assert not check_memory(p).ok  # 0,8 * 800 = 640
    assert check_memory(p, margin=1.0).ok


def test_memoria_desconocida(tmp_path, monkeypatch):
    monkeypatch.setattr(pu, "available_memory", lambda: None)
    r = check_memory(_archivo(tmp_path, "g.xlsx", 5000 * MB))
    assert r.ok
    assert r.available is None


def test_mensaje_en_mb(tmp_path, monkeypatch):
    p = _archivo(tmp_path, "m.csv", 100 * MB)  # necesita 714 MB
    monkeypatch.setattr(pu, "available_memory", lambda: 500 * MB)
    r = check_memory(p)
    assert not r.ok
    assert "714,0 MB" in r.message
    assert "500,0 MB libres" in r.message


def test_archivo_chico(tmp_path, monkeypatch):
    monkeypatch.setattr(pu, "available_memory", lambda: 4 * GB)
    p = tmp_path / "c.csv"
    p.write_bytes(b"a;b\n1;2\n")
    r = check_memory(p)
    assert r.ok
    assert r.message == ""
    assert BASE_BYTES <= r.needed < BASE_BYTES + 1000


def test_archivo_chico_xlsx(tmp_path, monkeypatch):
    import pandas as pd

    monkeypatch.setattr(pu, "available_memory", lambda: 4 * GB)
    p = tmp_path / "c.xlsx"
    pd.DataFrame({"a": ["1"]}).to_excel(p, index=False)
    assert check_memory(p).ok


def test_modulo_importable():
    assert memory.FACTOR_CSV == 6

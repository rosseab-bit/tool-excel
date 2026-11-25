import time
from pathlib import Path

import pandas as pd
import pytest

from excel_tool.core.errors import ArchivoBloqueado, ArchivoIlegible, ArchivoNoEncontrado
from excel_tool.core.io import CsvOptions, ReadOptions, detect_csv_options
from excel_tool.core.io.csv import read_csv

FILAS = [
    ["Código", "Nombre", "Localidad", "Importe"],
    ["00123", "Ñandú Pérez", "Córdoba", "1.234,56"],
    ["00124", "José Muñoz", "Rosario", "250,75"],
    ["NA", "null", "", "0,50"],
]


def _escribir(path: Path, filas, sep=";", encoding="utf-8-sig", fin="\n") -> Path:
    texto = fin.join(sep.join(f) for f in filas) + fin
    path.write_bytes(texto.encode(encoding))
    return path


@pytest.mark.parametrize("sep", [";", ",", "\t", "|"])
@pytest.mark.parametrize("encoding", ["utf-8-sig", "utf-8", "cp1252"])
def test_separadores_y_encodings(tmp_path, sep, encoding):
    filas = [[c.replace(",", ".") if sep == "," else c for c in f] for f in FILAS]
    p = _escribir(tmp_path / "a.csv", filas, sep, encoding)
    opts = detect_csv_options(p)
    assert opts == CsvOptions(separator=sep, encoding=encoding)
    r = read_csv(p)
    assert r.csv_options == opts
    assert list(r.df.columns) == FILAS[0]
    assert r.df.iloc[0].tolist() == filas[1]  # tildes, ñ y "00123" intactos
    assert r.header_row == 0


def test_texto_sin_conversion_de_nulos(tmp_path):
    r = read_csv(_escribir(tmp_path / "a.csv", FILAS))
    assert r.df.iloc[2].tolist() == ["NA", "null", "", "0,50"]
    assert all(isinstance(t, pd.StringDtype) and t.storage == "pyarrow" for t in r.df.dtypes)
    assert r.df.isna().sum().sum() == 0


def test_con_fixtures_csv(fixtures_generados):
    base = fixtures_generados / "csv"
    esperado = {
        "coma_utf8.csv": CsvOptions(",", "utf-8"),
        "con_bom.csv": CsvOptions(",", "utf-8-sig"),
        "excel_ar.csv": CsvOptions(";", "cp1252"),
        "tabulado.txt": CsvOptions("\t", "utf-8"),
    }
    for nombre, opts in esperado.items():
        r = read_csv(base / nombre)
        assert r.csv_options == opts, nombre
        assert list(r.df.columns) == ["Código", "Nombre", "Localidad", "Importe"]
        assert len(r.df) == 12
        assert r.df["Código"].str.startswith("0").any()  # ceros a la izquierda


def test_encabezado_en_fila_4(fixtures_generados, fixtures_esperados):
    import json

    esperado = json.loads((fixtures_esperados / "dificiles.json").read_text(encoding="utf-8"))
    info = esperado["con_basura.csv"]
    r = read_csv(fixtures_generados / "dificiles" / "con_basura.csv")
    assert r.header_row == info["fila_encabezado"] - 1
    assert list(r.df.columns) == info["columnas"]
    assert len(r.df) == info["filas_de_datos"]
    assert r.df.iloc[0, 0] == "00123"


def test_header_row_explicito_descarta_filas_de_arriba(fixtures_generados):
    p = fixtures_generados / "dificiles" / "con_basura.csv"
    r = read_csv(p, ReadOptions(header_row=3))
    assert list(r.df.columns) == ["Código", "Nombre", "Importe"]
    r0 = read_csv(p, ReadOptions(header_row=0))
    assert r0.header_row == 0
    assert len(r0.df) == 9


def test_nrows(fixtures_generados):
    r = read_csv(fixtures_generados / "explorar" / "para_explorar.csv", ReadOptions(nrows=5))
    assert len(r.df) == 5
    r = read_csv(fixtures_generados / "dificiles" / "con_basura.csv", ReadOptions(nrows=2))
    assert len(r.df) == 2
    assert list(r.df.columns) == ["Código", "Nombre", "Importe"]


def test_columnas_sin_nombre_y_repetidas(tmp_path):
    p = _escribir(tmp_path / "a.csv", [["Id", "", "Id", "", "Id"], ["1", "2", "3", "4", "5"]])
    r = read_csv(p)
    assert list(r.df.columns) == ["Id", "Columna B", "Id (2)", "Columna D", "Id (3)"]
    assert r.renamed_columns["Columna B"] == ""
    assert r.renamed_columns["Id (2)"] == "Id"
    assert len(r.warnings) == 2
    assert r.df.iloc[0].tolist() == ["1", "2", "3", "4", "5"]


def test_descarta_filas_vacias_al_final_no_las_del_medio(tmp_path):
    p = _escribir(
        tmp_path / "a.csv", [["A", "B"], ["1", "2"], ["", ""], ["3", "4"], ["", ""], ["", ""]]
    )
    r = read_csv(p)
    assert r.df.values.tolist() == [["1", "2"], ["", ""], ["3", "4"]]


def test_filas_con_mas_campos_que_la_muestra(tmp_path):
    filas = [["A", "B"]] + [["1", "2"]] * 3000 + [["3", "4", "5"]]
    r = read_csv(_escribir(tmp_path / "a.csv", filas))
    assert r.df.shape == (3001, 3)
    assert r.df.iloc[-1].tolist() == ["3", "4", "5"]
    assert r.df.columns[2] == "Columna C"


def test_campo_entre_comillas_con_separador_y_salto(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text('A;B\n"x;y";"línea1\nlínea2"\n', encoding="utf-8")
    r = read_csv(p)
    assert r.df.iloc[0].tolist() == ["x;y", "línea1\nlínea2"]


def test_encoding_invalido_mas_adelante_reintenta_cp1252(tmp_path):
    p = tmp_path / "a.csv"
    cuerpo = "A;B\n" + "x;y\n" * 100000 + "ñ;ú\n"
    p.write_bytes(cuerpo.encode("cp1252"))
    # los primeros 256 KB son ASCII: se detecta utf-8 y recién al final falla
    assert detect_csv_options(p).encoding == "utf-8"
    r = read_csv(p)
    assert r.csv_options.encoding == "cp1252"
    assert r.df.iloc[-1].tolist() == ["ñ", "ú"]


def test_opciones_explicitas(tmp_path):
    p = _escribir(tmp_path / "a.csv", FILAS, "|", "cp1252")
    r = read_csv(p, ReadOptions(csv=CsvOptions("|", "cp1252")))
    assert r.df.iloc[0, 1] == "Ñandú Pérez"
    with pytest.raises(ArchivoIlegible):
        read_csv(p, ReadOptions(csv=CsvOptions("|", "utf-8")))


def test_archivo_vacio(tmp_path):
    p = tmp_path / "vacio.csv"
    p.write_bytes(b"")
    with pytest.raises(ArchivoIlegible):
        detect_csv_options(p)
    with pytest.raises(ArchivoIlegible):
        read_csv(p)
    p.write_bytes(b"\n\n  \n")
    with pytest.raises(ArchivoIlegible):
        read_csv(p)


def test_inexistente(tmp_path):
    with pytest.raises(ArchivoNoEncontrado):
        read_csv(tmp_path / "no_esta.csv")
    with pytest.raises(ArchivoNoEncontrado):
        read_csv(tmp_path / "no_esta.csv", ReadOptions(csv=CsvOptions(";", "utf-8")))


def test_bloqueado(tmp_path, monkeypatch):
    p = _escribir(tmp_path / "a.csv", FILAS)
    original = Path.open

    def abrir(self, *a, **k):
        if self == p:
            raise PermissionError(13, "Permission denied")
        return original(self, *a, **k)

    monkeypatch.setattr(Path, "open", abrir)
    with pytest.raises(ArchivoBloqueado) as e:
        read_csv(p)
    assert "Cerrá el archivo" in e.value.mensaje
    assert isinstance(e.value.__cause__, PermissionError)


def test_bloqueado_en_lectura_completa(tmp_path, monkeypatch):
    import pandas as pd

    p = _escribir(tmp_path / "a.csv", FILAS)

    def falla(*a, **k):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(pd, "read_csv", falla)
    with pytest.raises(ArchivoBloqueado):
        read_csv(p)


def test_rendimiento_2M(fixtures_grandes):
    p = fixtures_grandes / "rendimiento_2M.csv"
    t = time.perf_counter()
    r = read_csv(p)
    seg = time.perf_counter() - t
    print(f"\nrendimiento_2M.csv: {len(r.df)} filas en {seg:.1f} s")
    assert len(r.df) == 2_000_000
    assert seg <= 20

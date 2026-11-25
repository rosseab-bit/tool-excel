"""Verifica que los archivos de prueba generados tengan el contenido esperado."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

CLIENTES = [
    ("00123", "Ñandú Pérez", "Córdoba"),
    ("00124", "José Muñoz", "Neuquén"),
    ("00125", "María Núñez", "Tucumán"),
    ("00126", "Sebastián Peña", "Rosario"),
    ("00127", "Inés Ibáñez", "Mendoza"),
    ("00128", "Ángel Güemes", "Salta"),
    ("00129", "Lucía Fernández", "Bahía Blanca"),
    ("01000", "Cañuelas S.A.", "Cañuelas"),
    ("01001", "Dueñas y Cía", "Río Cuarto"),
    ("01002", "Óscar Sánchez", "Paraná"),
    ("01003", "Úrsula Vázquez", "Maipú"),
    ("01004", "Zoé Rodríguez", "Santa Fe"),
]
SALDOS = [
    "1234.56", "250.75", "98765.43", "0.5", "15000.25", "310.1",
    "4321.99", "777.77", "12.34", "56789.01", "1000.5", "2.25",
]  # fmt: skip
IMPORTES_AR = [
    "1.234,56", "250,75", "98.765,43", "0,50", "15.000,25", "310,10",
    "4.321,99", "777,77", "12,34", "56.789,01", "1.000,50", "2,25",
]  # fmt: skip
IMPORTES_PUNTO = [f"{float(s):.2f}" for s in SALDOS]

# (archivo, separador, encoding, importes esperados)
CSVS = [
    ("csv/excel_ar.csv", ";", "cp1252", IMPORTES_AR),
    ("csv/coma_utf8.csv", ",", "utf-8", IMPORTES_PUNTO),
    ("csv/con_bom.csv", ",", "utf-8-sig", IMPORTES_PUNTO),
    ("csv/tabulado.txt", "\t", "utf-8", IMPORTES_PUNTO),
]


def _leer_excel(ruta: Path, hoja=0) -> pd.DataFrame:
    return pd.read_excel(ruta, sheet_name=hoja, dtype=str, engine="calamine")


@pytest.mark.parametrize(("archivo", "sep", "enc", "importes"), CSVS)
def test_csv_conserva_tildes_enies_y_ceros(fixtures_generados, archivo, sep, enc, importes):
    df = pd.read_csv(fixtures_generados / archivo, sep=sep, encoding=enc, dtype=str)
    assert list(df.columns) == ["Código", "Nombre", "Localidad", "Importe"]
    assert list(df.itertuples(index=False, name=None))[0][:3] == CLIENTES[0]
    assert [tuple(f) for f in df[["Código", "Nombre", "Localidad"]].to_numpy()] == CLIENTES
    assert list(df["Importe"]) == importes
    assert "00123" in set(df["Código"])


def test_con_bom_empieza_con_bom(fixtures_generados):
    assert (fixtures_generados / "csv/con_bom.csv").read_bytes().startswith(b"\xef\xbb\xbf")
    assert not (fixtures_generados / "csv/coma_utf8.csv").read_bytes().startswith(b"\xef\xbb\xbf")


def test_excel_ar_es_cp1252_y_no_utf8(fixtures_generados):
    datos = (fixtures_generados / "csv/excel_ar.csv").read_bytes()
    with pytest.raises(UnicodeDecodeError):
        datos.decode("utf-8")
    assert "Ñandú".encode("cp1252") in datos


def test_xlsx_varias_hojas(fixtures_generados):
    ruta = fixtures_generados / "formatos/varias_hojas.xlsx"
    hojas = pd.read_excel(ruta, sheet_name=None, dtype=str, engine="calamine")
    assert list(hojas) == ["Clientes", "Ventas", "Resumen año"]
    assert [h.shape for h in hojas.values()] == [(12, 4), (25, 6), (2, 2)]
    clientes = hojas["Clientes"]
    assert [tuple(f) for f in clientes[["Código", "Nombre", "Localidad"]].to_numpy()] == CLIENTES
    assert list(clientes["Saldo"]) == SALDOS
    assert "Observación" in hojas["Ventas"].columns
    assert set(hojas["Ventas"]["Código"]) <= {c[0] for c in CLIENTES}


@pytest.mark.parametrize("archivo", ["formatos/chico.xls", "formatos/chico.xlsb"])
def test_xls_y_xlsb_valores_exactos(fixtures_generados, archivo):
    df = _leer_excel(fixtures_generados / archivo)
    assert list(df.columns) == ["Código", "Nombre", "Localidad", "Saldo"]
    assert [tuple(f) for f in df[["Código", "Nombre", "Localidad"]].to_numpy()] == CLIENTES
    assert list(df["Saldo"]) == SALDOS


def test_xls_y_xlsb_coinciden_con_la_hoja_del_xlsx(fixtures_generados):
    base = _leer_excel(fixtures_generados / "formatos/varias_hojas.xlsx", "Clientes")
    for archivo in ("formatos/chico.xls", "formatos/chico.xlsb"):
        pd.testing.assert_frame_equal(_leer_excel(fixtures_generados / archivo), base)


def test_xlsb_es_un_zip_con_las_partes_esperadas(fixtures_generados):
    import zipfile

    with zipfile.ZipFile(fixtures_generados / "formatos/chico.xlsb") as z:
        assert set(z.namelist()) == {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.bin",
            "xl/_rels/workbook.bin.rels",
            "xl/worksheets/sheet1.bin",
            "xl/styles.bin",
        }


def test_para_explorar_csv_y_xlsx_tienen_los_mismos_datos(fixtures_generados):
    csv = pd.read_csv(
        fixtures_generados / "explorar/para_explorar.csv", sep=";", encoding="utf-8-sig", dtype=str
    )
    xlsx = _leer_excel(fixtures_generados / "explorar/para_explorar.xlsx")
    assert csv.shape == xlsx.shape == (200, 13)
    pd.testing.assert_frame_equal(csv.fillna(""), xlsx.fillna(""))
    codigos = csv["Código"].dropna()
    assert codigos.str.startswith("0").all()
    assert codigos.str.len().eq(5).all()
    assert csv["Descripción"].str.contains("Ñandú|Cañería|Lámpara|Camión|Árbol").any()
    assert {"Norte", "NORTE", "Norte "} <= set(csv["Región"].dropna())


def test_perfil_esperado_coincide_con_los_datos(fixtures_generados, fixtures_esperados):
    """Chequeo grueso del perfil.json contra el archivo (la verificación fina es de la Fase 2)."""
    perfil = json.loads((fixtures_esperados / "perfil.json").read_text(encoding="utf-8"))
    assert "_nota" in perfil
    df = pd.read_csv(
        fixtures_generados / "explorar/para_explorar.csv",
        sep=";",
        encoding="utf-8-sig",
        dtype=str,
        keep_default_na=False,
    )
    assert perfil["filas"] == len(df) == 200
    assert perfil["columnas"] == df.shape[1] == len(perfil["columnas_detalle"])
    vacias = (df == "").all(axis=1)
    assert perfil["filas_vacias"] == int(vacias.sum())
    assert perfil["filas_repetidas"] == int(df[~vacias].duplicated().sum())
    for pos, col in enumerate(perfil["columnas_detalle"]):
        serie = df.iloc[:, pos]
        assert col["posicion"] == pos + 1
        assert col["vacios"] == int((serie == "").sum())
        assert col["distintos"] == serie[serie != ""].nunique()
    tipos = {c["nombre"]: c["tipo"] for c in perfil["columnas_detalle"]}
    assert tipos["Mixta"] == "mixto"
    assert tipos["Vacía"] == "vacia"
    mixta = next(c for c in perfil["columnas_detalle"] if c["nombre"] == "Mixta")
    assert set(mixta["desglose_tipos"]) == {"entero", "texto"}


def test_generar_dos_veces_da_lo_mismo(generador_fixtures, tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    generador_fixtures.generar(a, sin_grandes=True)
    generador_fixtures.generar(b, sin_grandes=True)
    archivos = sorted(p.relative_to(a) for p in a.rglob("*") if p.is_file())
    assert archivos == sorted(p.relative_to(b) for p in b.rglob("*") if p.is_file())
    assert len(archivos) >= 9
    for rel in archivos:
        if rel.suffix in {".csv", ".txt", ".json", ""} or rel.name == "danado.xlsx":
            assert (a / rel).read_bytes() == (b / rel).read_bytes(), rel
        else:
            if rel.suffix == ".xlsx":
                ha = pd.read_excel(a / rel, sheet_name=None, dtype=str, engine="calamine")
                hb = pd.read_excel(b / rel, sheet_name=None, dtype=str, engine="calamine")
                assert list(ha) == list(hb)
                for nombre in ha:
                    pd.testing.assert_frame_equal(ha[nombre], hb[nombre])
            else:
                pd.testing.assert_frame_equal(_leer_excel(a / rel), _leer_excel(b / rel))


def test_generar_no_escribe_fuera_de_la_salida(generador_fixtures, tmp_path):
    antes = (generador_fixtures.ESPERADOS / "perfil.json").read_bytes()
    generador_fixtures.generar(tmp_path / "x", sin_grandes=True)
    assert (generador_fixtures.ESPERADOS / "perfil.json").read_bytes() == antes
    assert (tmp_path / "x" / "esperados" / "perfil.json").is_file()
    with pytest.raises(ValueError):
        generador_fixtures.ruta_segura(tmp_path / "x", "../afuera.txt")


def test_fixtures_grandes_se_saltean_si_faltan(request, fixtures_generados):
    marca = fixtures_generados / "grandes" / "LISTO"
    if marca.is_file():
        pytest.skip("Los archivos grandes existen; este test solo cubre el caso contrario.")
    with pytest.raises(pytest.skip.Exception):
        request.getfixturevalue("fixtures_grandes")

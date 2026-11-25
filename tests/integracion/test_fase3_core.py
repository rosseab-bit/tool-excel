"""Fase 3, core de punta a punta: leer con `read_table`, medir con `limits` y dividir con `split`.

Los tests con archivos grandes se saltean (con aviso) si faltan: se generan con
`.venv/bin/python scripts/make_fixtures.py` (sin --sin-grandes). Los números esperados están
calculados a mano en cada test o salen de `tests/fixtures/esperados/fase3_*.json`, que a su vez
salen de la definición de los datos del generador (no del código que se prueba).
"""

from __future__ import annotations

import json
import math
import re
from collections import Counter
from pathlib import Path

import pytest

from excel_tool import platform_utils
from excel_tool.core import limits
from excel_tool.core.io import read_table
from excel_tool.core.limits import SizeEstimate
from excel_tool.core.split import SplitSpec, plan_split, split_dataframe

# Valores de PLAN §1.2 (limites_default.json), a mano.
EXCEL_MAX_FILAS_DATOS = 1_048_576 - 1  # el encabezado ocupa una fila
EXCEL_COMODO = 300_000
EXCEL_COMODO_MB = 50


@pytest.fixture(autouse=True)
def _datos_de_usuario_aislados(tmp_path, monkeypatch):
    """Que un `limites.json` del usuario que corre los tests no cambie los resultados."""
    carpeta = tmp_path / "datos_usuario"
    carpeta.mkdir()
    monkeypatch.setattr(platform_utils, "user_data_dir", lambda: carpeta)
    return carpeta


def _esperado(fixtures_esperados: Path, nombre: str) -> dict:
    return json.loads((fixtures_esperados / nombre).read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def tabla_400k(request):
    grandes = request.getfixturevalue("fixtures_grandes")
    return read_table(grandes / "fase3_400k.csv").df


@pytest.fixture(scope="module")
def tabla_1600k(request):
    grandes = request.getfixturevalue("fixtures_grandes")
    return read_table(grandes / "fase3_dividir_1600k.csv").df


@pytest.fixture(scope="module")
def tabla_nombres(fixtures_generados):
    return read_table(fixtures_generados / "fase3" / "nombres_raros.csv").df


# ------------------------------------------------------------------------------ límites por default


def test_limites_por_defecto_son_los_del_plan():
    lim = limits.load_limits()
    assert lim.excel_max_rows == 1_048_576
    assert lim.excel_comfortable_rows == EXCEL_COMODO
    assert lim.excel_comfortable_mb == EXCEL_COMODO_MB
    assert lim.gs_max_cells == 20_000_000


# ------------------------------------------------------------------------------ semáforo


def test_escenario_grande_es_rojo_para_excel(tabla_1600k, fixtures_esperados):
    esp = _esperado(fixtures_esperados, "fase3_grandes.json")
    total = esp["filas"]["fase3_dividir_1600k.csv"]
    assert len(tabla_1600k) == total == 1_600_000
    est = limits.estimate(tabla_1600k)
    assert (est.rows, est.columns, est.cells) == (1_600_000, 4, 1_600_001 * 4)

    chk = limits.check(est, "excel", "xlsx")
    assert chk.level == "obligatorio"
    assert chk.reasons  # explica por qué, en castellano
    # A mano: por filas cómodas, ceil(1.600.000 / 300.000) = 6 partes. Si los MB cómodos piden más
    # partes, manda el mayor de los dos.
    por_filas = math.ceil(1_600_000 / EXCEL_COMODO)
    assert por_filas == 6
    por_mb = math.ceil(est.approx_mb_xlsx / EXCEL_COMODO_MB)
    assert chk.suggested_parts == max(por_filas, por_mb)
    assert chk.suggested_parts >= 6
    assert chk.suggested_rows_per_part == math.ceil(1_600_000 / chk.suggested_parts)
    assert chk.suggested_rows_per_part <= EXCEL_COMODO

    # "Solo guardar" en CSV no tiene límite; en xlsx rige el de la hoja: ceil(1,6M / 1.048.575) = 2.
    assert limits.check(est, "solo_guardar", "csv").level == "ok"
    duro = limits.check(est, "solo_guardar", "xlsx")
    assert duro.level == "obligatorio"
    assert duro.suggested_parts == math.ceil(1_600_000 / EXCEL_MAX_FILAS_DATOS) == 2


def test_400k_amarillo_para_excel_y_verde_para_solo_guardar(tabla_400k):
    assert len(tabla_400k) == 400_000
    est = limits.estimate(tabla_400k)
    chk = limits.check(est, "excel", "xlsx")
    assert chk.level == "recomendado"
    assert chk.suggested_parts >= math.ceil(400_000 / EXCEL_COMODO) == 2
    assert limits.check(est, "solo_guardar", "csv").level == "ok"
    assert limits.check(est, "solo_guardar", "xlsx").level == "ok"  # 400k entra en una hoja
    # Una sola hoja alcanza: no es rojo.
    assert limits.check(est, "excel", "csv").level == "recomendado"


def test_limite_de_limites_json_cambia_la_recomendacion(_datos_de_usuario_aislados):
    # Resultado de 400k filas x 4 columnas, armado a mano (no hace falta el archivo grande).
    est = SizeEstimate(
        rows=400_000, columns=4, cells=400_001 * 4, approx_mb_xlsx=15, approx_mb_csv=14
    )
    assert limits.check(est, "excel", "xlsx").level == "recomendado"

    archivo = _datos_de_usuario_aislados / "limites.json"
    archivo.write_text(json.dumps({"excel": {"comfortable_rows": 500_000}}), encoding="utf-8")
    lim = limits.load_limits()
    assert lim.excel_comfortable_rows == 500_000
    assert limits.check(est, "excel", "xlsx").level == "ok"

    archivo.write_text(json.dumps({"excel": {"max_rows": 350_000}}), encoding="utf-8")
    chk = limits.check(est, "excel", "xlsx")
    assert chk.level == "obligatorio"
    assert chk.suggested_parts >= 2

    # Un archivo roto no rompe nada: vuelven los valores por defecto.
    archivo.write_text("esto no es json", encoding="utf-8")
    assert limits.load_limits().excel_comfortable_rows == EXCEL_COMODO
    assert limits.check(est, "excel", "xlsx").level == "recomendado"


def test_limite_de_limites_json_con_el_archivo_leido(tabla_400k, _datos_de_usuario_aislados):
    est = limits.estimate(tabla_400k)
    assert limits.check(est, "excel", "xlsx").level == "recomendado"
    (_datos_de_usuario_aislados / "limites.json").write_text(
        json.dumps({"excel": {"comfortable_rows": 450_000, "comfortable_mb": 500}}),
        encoding="utf-8",
    )
    assert limits.check(est, "excel", "xlsx").level == "ok"


# ------------------------------------------------------------------------------ dividir


def test_division_por_filas_ninguna_parte_supera_el_limite(tabla_1600k):
    total = len(tabla_1600k)
    # Con el máximo de la hoja: 1.048.575 + 551.425.
    spec = SplitSpec("rows")
    partes = list(
        split_dataframe(
            tabla_1600k, spec, base_name="Ventas", max_rows_per_part=EXCEL_MAX_FILAS_DATOS
        )
    )
    assert [len(df) for _, df in partes] == [
        EXCEL_MAX_FILAS_DATOS,
        1_600_000 - EXCEL_MAX_FILAS_DATOS,
    ]
    assert all(p.rows == len(df) <= EXCEL_MAX_FILAS_DATOS for p, df in partes)
    assert sum(len(df) for _, df in partes) == total
    # El orden de las filas se mantiene: los IDs de las partes, juntos, son 1..N.
    ids = [int(v) for _, df in partes for v in (df["ID"].iloc[0], df["ID"].iloc[-1])]
    assert ids == [1, EXCEL_MAX_FILAS_DATOS, EXCEL_MAX_FILAS_DATOS + 1, total]

    # Con las partes sugeridas por el semáforo (cómodo): cada una dentro de lo cómodo.
    chk = limits.check(limits.estimate(tabla_1600k), "excel", "xlsx")
    spec = SplitSpec("rows", rows_per_part=chk.suggested_rows_per_part)
    planes = plan_split(tabla_1600k, spec, base_name="Ventas")
    assert len(planes) == chk.suggested_parts
    assert all(p.rows <= EXCEL_COMODO for p in planes)
    assert sum(p.rows for p in planes) == total


def test_rows_keep_groups_no_corta_ningun_cliente(tabla_1600k):
    cuentas = Counter(tabla_1600k["Cliente"])
    assert len(cuentas) > 1000
    assert len(set(cuentas.values())) > 100  # tamaños de grupo muy variados
    assert max(cuentas.values()) < EXCEL_COMODO  # ningún grupo pasa solo del límite
    # Hay clientes no contiguos en el archivo (si no, el test no prueba nada).
    clientes = tabla_1600k["Cliente"]
    saltos = (clientes != clientes.shift()).sum()
    assert saltos > len(cuentas)

    spec = SplitSpec("rows_keep_groups", rows_per_part=EXCEL_COMODO, group_column="Cliente")
    vistos: set[str] = set()
    contados: Counter = Counter()
    total = 0
    cantidad = 0
    for plan, parte in split_dataframe(tabla_1600k, spec, base_name="Ventas"):
        cantidad += 1
        assert len(parte) == plan.rows <= EXCEL_COMODO
        en_parte = Counter(parte["Cliente"])
        assert not (set(en_parte) & vistos), "un cliente quedó en dos partes"
        vistos |= set(en_parte)
        contados += en_parte
        total += len(parte)
    assert cantidad >= math.ceil(1_600_000 / EXCEL_COMODO)
    assert total == len(tabla_1600k)
    assert contados == cuentas  # cada cliente completo, con todas sus filas


def test_by_value_subdivide_la_region_grande(tabla_1600k, fixtures_esperados):
    esp = _esperado(fixtures_esperados, "fase3_grandes.json")["regiones_dividir"]
    assert esp["Norte"] == 1_200_000 > EXCEL_MAX_FILAS_DATOS
    spec = SplitSpec("by_value", value_columns=["Región"])
    por_region: dict[str, list[int]] = {}
    nombres: list[str] = []
    total = 0
    for plan, parte in split_dataframe(
        tabla_1600k, spec, base_name="Ventas", max_rows_per_part=EXCEL_MAX_FILAS_DATOS
    ):
        assert len(parte) == plan.rows <= EXCEL_MAX_FILAS_DATOS
        valores = set(parte["Región"])
        assert len(valores) == 1, "una parte mezcla regiones"
        assert plan.value == (next(iter(valores)),)
        por_region.setdefault(next(iter(valores)), []).append(len(parte))
        nombres.append(plan.name)
        total += len(parte)
    # Norte: 1.200.000 = 1.048.575 + 151.425; las demás, una parte cada una.
    assert sorted(por_region["Norte"]) == [151_425, EXCEL_MAX_FILAS_DATOS]
    for r in ("Sur", "Centro", "Cuyo", "Litoral"):
        assert por_region[r] == [100_000]
    assert {r: sum(v) for r, v in por_region.items()} == esp
    assert len(nombres) == 6
    assert len({n.lower() for n in nombres}) == 6
    assert total == len(tabla_1600k)


def test_by_value_sin_limite_da_una_parte_por_region(tabla_1600k, fixtures_esperados):
    esp = _esperado(fixtures_esperados, "fase3_grandes.json")["regiones_dividir"]
    planes = plan_split(
        tabla_1600k, SplitSpec("by_value", value_columns=["Región"]), base_name="Ventas"
    )
    assert {p.value[0]: p.rows for p in planes} == esp
    assert sum(p.rows for p in planes) == len(tabla_1600k)


def test_400k_se_divide_en_dos_con_las_partes_sugeridas(tabla_400k):
    chk = limits.check(limits.estimate(tabla_400k), "excel", "xlsx")
    planes = plan_split(
        tabla_400k, SplitSpec("rows", rows_per_part=chk.suggested_rows_per_part), base_name="R"
    )
    assert len(planes) == chk.suggested_parts
    assert sum(p.rows for p in planes) == 400_000
    assert all(p.rows <= EXCEL_COMODO for p in planes)


# ------------------------------------------------------------------------------ columnas sugeridas


def test_suggest_split_columns_propone_region_y_no_id_grande(tabla_1600k):
    sug = limits.suggest_split_columns(tabla_1600k, "excel", "xlsx")
    nombres = [s.column for s in sug]
    assert "Región" in nombres
    assert "ID" not in nombres
    assert "Cliente" not in nombres  # 5.000 valores: demasiados archivos
    region = next(s for s in sug if s.column == "Región")
    assert region.parts == 5
    assert region.largest_part_rows == 1_200_000
    assert region.fits is False  # Norte no entra en lo cómodo de Excel


def test_suggest_split_columns_propone_sucursal_y_no_id_chico(tabla_nombres):
    sug = limits.suggest_split_columns(tabla_nombres, "excel", "xlsx")
    nombres = [s.column for s in sug]
    assert "Sucursal" in nombres
    assert "ID" not in nombres
    sucursal = next(s for s in sug if s.column == "Sucursal")
    assert sucursal.fits is True  # 100 filas por parte como mucho


# ------------------------------------------------------------------------------ nombres de archivo

_PROHIBIDOS = set('<>:"/\\|?*')
_RESERVADOS = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)}
_RESERVADOS |= {f"LPT{i}" for i in range(1, 10)}


def _es_valido_en_windows(nombre_con_extension: str) -> str | None:
    """None si es válido; si no, el motivo. Reglas escritas acá, sin usar el código del producto."""
    if not nombre_con_extension:
        return "vacío"
    if any(c in _PROHIBIDOS or ord(c) < 32 for c in nombre_con_extension):
        return "caracter prohibido"
    if nombre_con_extension.endswith((" ", ".")):
        return "termina en espacio o punto"
    if nombre_con_extension.split(".")[0].upper() in _RESERVADOS:
        return "nombre reservado"
    if len(nombre_con_extension) > 255:
        return "demasiado largo"
    return None


@pytest.mark.parametrize("patron", [None, "{valor}", "{nombre}-{valor}.v1"])
def test_nombres_raros_dan_nombres_validos_y_sin_repetir(tabla_nombres, fixtures_esperados, patron):
    esp = _esperado(fixtures_esperados, "fase3_nombres.json")
    assert len(tabla_nombres) == esp["filas"]
    spec = SplitSpec("by_value", value_columns=["Sucursal"], name_pattern=patron)
    total = 0
    nombres: list[str] = []
    filas_por_valor: dict[str, int] = {}
    for plan, parte in split_dataframe(tabla_nombres, spec, base_name="Ventas"):
        for ext in (".csv", ".xlsx"):
            motivo = _es_valido_en_windows(plan.name + ext)
            assert motivo is None, f"{plan.name!r}{ext}: {motivo}"
        nombres.append(plan.name)
        total += len(parte)
        filas_por_valor[plan.value[0]] = len(parte)
    assert total == esp["filas"]
    # Sin repetir, tampoco distinguiendo mayúsculas (Windows no las distingue).
    assert len({n.lower() for n in nombres}) == len(nombres)
    # Los valores se agrupan tal cual (distinguiendo mayúsculas): una parte por valor.
    assert filas_por_valor == esp["filas_por_valor"]
    assert len(nombres) == esp["valores_distintos"]


def test_nombres_raros_tambien_como_hojas(tabla_nombres):
    spec = SplitSpec("sheets", rows_per_part=20)
    nombres = [p.name for p in plan_split(tabla_nombres, spec, base_name="Ventas")]
    assert nombres
    assert len({n.lower() for n in nombres}) == len(nombres)
    for n in nombres:
        assert 1 <= len(n) <= 31
        assert not (set(n) & set("[]:*?/\\")), n


def test_suma_de_filas_chico_en_todos_los_modos(tabla_nombres):
    total = len(tabla_nombres)
    for spec in (
        SplitSpec("rows", rows_per_part=17),
        SplitSpec("rows_keep_groups", rows_per_part=25, group_column="Sucursal"),
        SplitSpec("by_value", value_columns=["Sucursal"]),
        SplitSpec("sheets", rows_per_part=30),
    ):
        partes = list(split_dataframe(tabla_nombres, spec, base_name="Ventas"))
        assert sum(len(df) for _, df in partes) == total, spec.mode
        assert [p.rows for p, _ in partes] == [len(df) for _, df in partes]
        assert list(tabla_nombres.columns) == list(partes[0][1].columns)


def test_valor_vacio_y_mayusculas_en_by_value(tabla_nombres):
    planes = plan_split(
        tabla_nombres, SplitSpec("by_value", value_columns=["Sucursal"]), base_name="Ventas"
    )
    por_valor = {p.value[0]: p for p in planes}
    assert por_valor[""].rows == 6
    assert re.search(r"vacio", por_valor[""].name, re.IGNORECASE)
    assert "Sur" in por_valor and "sur" in por_valor
    assert por_valor["Sur"].name.lower() != por_valor["sur"].name.lower()

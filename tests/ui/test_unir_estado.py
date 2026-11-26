"""Tests sin pantalla de EstadoUnion (F5-10)."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import ReadOptions, read_table
from excel_tool.core.join import SourceSpec, run_join
from excel_tool.core.recipe import load_recipe, save_recipe
from excel_tool.ui.unir.estado import EstadoUnion, Fuente, FuenteSecundaria, leer_fuente

SUMA = AggregationSpec([GroupBy("Nro Cliente")], [Aggregation("Monto", "sum", "Total cobrado")])


@pytest.fixture
def carpeta(fixtures_generados, tmp_path) -> Path:
    destino = tmp_path / "datos"
    shutil.copytree(fixtures_generados / "escenario_3", destino)
    return destino


def _fuente(carpeta: Path, nombre: str, columnas: list[str]) -> Fuente:
    spec = SourceSpec(nombre, carpeta / nombre, ReadOptions(), columnas)
    return Fuente(spec, read_table(spec.path).df)


def _sec(carpeta: Path, nombre: str, pares, columnas, **kw) -> FuenteSecundaria:
    base = _fuente(carpeta, nombre, columnas)
    return FuenteSecundaria(base.spec, base.df, key_pairs=pares, **kw)


def _estado(carpeta: Path, modo_ventas="detail") -> EstadoUnion:
    return EstadoUnion(
        principal=_fuente(carpeta, "clientes.xlsx", ["ID cliente", "Nombre", "Región"]),
        claves=["ID cliente"],
        secundarios=[
            _sec(
                carpeta,
                "ventas.csv",
                [("ID cliente", "ID cliente")],
                ["Fecha", "Importe"],
                mode=modo_ventas,
            ),
            _sec(
                carpeta,
                "cobranzas.xlsx",
                [("ID cliente", "Nro Cliente")],
                [],
                mode="summary",
                summary=SUMA,
            ),
        ],
    )


def test_plan_a_mano_lo_ejecuta_el_core(carpeta):
    estado = _estado(carpeta)
    plan = estado.plan()
    assert [s.name for s in plan.secondaries] == ["ventas.csv", "cobranzas.xlsx"]
    assert plan.secondaries[1].summary == SUMA
    resultado = run_join(plan, estado.tablas())
    assert "Total cobrado" in resultado.df.columns
    assert len(resultado.df) >= len(estado.principal.df)


def test_secundario_nuevo_empieza_sin_modo(carpeta):
    sec = _sec(carpeta, "ventas.csv", [("ID cliente", "ID cliente")], [])
    assert sec.mode is None


def test_modo_sin_elegir_da_error_claro(carpeta):
    estado = _estado(carpeta, modo_ventas=None)
    with pytest.raises(ExcelToolError, match="ventas.csv.*detalle o un resumen"):
        estado.plan()


def test_faltan_datos_en_el_plan(carpeta):
    with pytest.raises(ExcelToolError, match="archivo principal"):
        EstadoUnion().plan()
    estado = _estado(carpeta)
    estado.secundarios[1].summary = None
    with pytest.raises(ExcelToolError, match="resumen de"):
        estado.plan()
    estado = _estado(carpeta)
    estado.secundarios.clear()
    with pytest.raises(ExcelToolError, match="sumar al menos"):
        estado.plan()


def test_columnas_del_estado_mandan_sobre_las_del_spec(carpeta):
    estado = _estado(carpeta)
    estado.columnas = {"ventas.csv": ["Importe"]}
    assert estado.plan().secondaries[0].columns == ["Importe"]


def _receta(carpeta: Path) -> Path:
    receta = carpeta / "mes.json"
    save_recipe(_estado(carpeta).plan(), receta)
    return receta


def test_desde_receta_con_todos_los_archivos(carpeta):
    carga = load_recipe(_receta(carpeta))
    estado = EstadoUnion.desde_receta(carga)
    assert estado.faltantes == []
    assert estado.claves == ["ID cliente"]
    assert estado.principal.perfil is not None and len(estado.principal.df) > 0
    assert [s.mode for s in estado.secundarios] == ["detail", "summary"]
    assert estado.plan() == carga.plan
    run_join(estado.plan(), estado.tablas())


def test_desde_receta_con_archivo_faltante(carpeta):
    receta = _receta(carpeta)
    (carpeta / "cobranzas.xlsx").unlink()
    estado = EstadoUnion.desde_receta(load_recipe(receta))
    assert estado.faltantes == ["cobranzas.xlsx"]
    assert estado.hay_datos
    with pytest.raises(ExcelToolError, match="cobranzas.xlsx"):
        estado.plan()
    with pytest.raises(ExcelToolError, match="cobranzas.xlsx"):
        estado.tablas()
    # el archivo nuevo del mes completa el hueco y se conserva lo elegido
    otro = carpeta / "cobranzas_nuevo.xlsx"
    shutil.copy(carpeta / "clientes.xlsx", otro)
    estado.completar("cobranzas.xlsx", leer_fuente(SourceSpec("x", otro, ReadOptions(), [])))
    assert estado.faltantes == []
    sec = estado.secundarios[1]
    assert sec.spec.path == otro and sec.spec.name == "cobranzas.xlsx"
    assert sec.mode == "summary" and sec.summary == SUMA


def test_completar_nombre_inexistente(carpeta):
    with pytest.raises(ExcelToolError):
        _estado(carpeta).completar("nada.csv", _fuente(carpeta, "ventas.csv", []))


def test_diagnosticar_reusa_el_resultado_si_el_plan_no_cambio(carpeta, monkeypatch):
    from excel_tool.ui.unir import estado as modulo

    llamadas = []
    real = modulo.diagnose_join

    def contado(plan, tablas, **kw):
        llamadas.append(1)
        return real(plan, tablas, **kw)

    monkeypatch.setattr(modulo, "diagnose_join", contado)
    est = _estado(carpeta)
    a = est.diagnosticar(est.plan(), est.tablas())
    b = est.diagnosticar(est.plan(), est.tablas())
    assert a is b and len(llamadas) == 1
    # el orden de las columnas no cambia filas ni tamaño: no vuelve a calcular
    est.orden = list(reversed(est.plan().principal.columns))
    est.diagnosticar(est.plan(), est.tablas())
    assert len(llamadas) == 1
    # un cambio real del plan sí recalcula
    est.secundarios[0].mode = "first"
    c = est.diagnosticar(est.plan(), est.tablas())
    assert len(llamadas) == 2 and c is not a
    est.olvidar_diagnosticos()
    est.diagnosticar(est.plan(), est.tablas())
    assert len(llamadas) == 3

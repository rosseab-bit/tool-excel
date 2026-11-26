"""Tests de core/recipe (F4-04): guardar y cargar recetas."""

from __future__ import annotations

import json
import shutil
from dataclasses import replace
from pathlib import Path

import pytest

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import CsvOptions, ReadOptions, WriteOptions, read_table, write_table
from excel_tool.core.join import JoinPlan, SecondarySpec, SourceSpec, run_join
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import OutputSpec, save_result
from excel_tool.core.recipe import RecipeLoad, load_recipe, save_recipe
from excel_tool.core.split import SplitSpec


@pytest.fixture(params=["csv", "xlsx"])
def carpeta(request, fixtures_generados, tmp_path):
    """Carpeta con los datos del escenario 3 (ventas en CSV o xlsx)."""
    datos = tmp_path / "datos"
    datos.mkdir()
    base = fixtures_generados / "escenario_3"
    shutil.copy(base / "clientes.xlsx", datos / "clientes.xlsx")
    shutil.copy(base / "cobranzas.xlsx", datos / "cobranzas.xlsx")
    if request.param == "csv":
        shutil.copy(base / "ventas.csv", datos / "ventas.csv")
    else:
        write_table(read_table(base / "ventas.csv").df, datos / "ventas.xlsx", WriteOptions("xlsx"))
    return datos


def _plan(datos: Path, salida: Path) -> JoinPlan:
    ventas = next(datos.glob("ventas.*"))
    ro = ReadOptions(csv=CsvOptions(";", "cp1252")) if ventas.suffix == ".csv" else ReadOptions()
    return JoinPlan(
        SourceSpec(
            "clientes",
            datos / "clientes.xlsx",
            ReadOptions(sheet=None, header_row=0),
            ["ID cliente", "Región"],
        ),
        KeyOptions(strip_leading_zeros=True),
        [
            SecondarySpec(
                "ventas",
                ventas,
                ro,
                [],
                [("ID cliente", "ID cliente")],
                mode="summary",
                summary=AggregationSpec(
                    [GroupBy("ID cliente")], [Aggregation("Importe", "sum", "Total comprado")]
                ),
                prefix="V",
            ),
            SecondarySpec(
                "cobranzas",
                datos / "cobranzas.xlsx",
                ReadOptions(),
                ["Fecha", "Monto"],
                [("ID cliente", "Nro Cliente")],
                how="left",
                mode="summary",
                summary=AggregationSpec(
                    [GroupBy("Nro Cliente"), GroupBy("Fecha", "month")],
                    [Aggregation("Monto", "sum", "Total cobrado")],
                ),
            ),
        ],
        output_column_order=None,
        final_summary=AggregationSpec(
            [GroupBy("Región")], [Aggregation("VTotal comprado", "sum", "Total vendido")]
        ),
        output=OutputSpec(
            salida / "resultado.xlsx",
            "xlsx",
            "excel",
            split=SplitSpec("by_value", value_columns=["Región"], name_pattern=None),
            sheet_name="Datos",
            column_types={"Región": "texto"},
        ),
    )


def _correr(plan: JoinPlan, carpeta_salida: Path):
    tablas = {
        s.name: read_table(s.path, s.read_options).df for s in [plan.principal, *plan.secondaries]
    }
    res = run_join(plan, tablas)
    out = replace(plan.output, path=carpeta_salida / plan.output.path.name)
    guardado = save_result(res.df, out)
    return res.df, guardado


def _contenido(guardado):
    return sorted((p.path.name, p.rows, p.value) for p in guardado.parts), guardado.rows_total


def test_ida_y_vuelta_mismo_plan_y_mismo_resultado(carpeta, tmp_path):
    plan = _plan(carpeta, tmp_path / "salida1")
    receta = carpeta / "mes.json"
    save_recipe(plan, receta)
    cargada = load_recipe(receta)
    assert isinstance(cargada, RecipeLoad)
    assert cargada.plan == plan
    assert cargada.missing == []
    assert not receta.with_name("mes.json.tmp").exists()

    (tmp_path / "salida1").mkdir()
    (tmp_path / "salida2").mkdir()
    df1, g1 = _correr(plan, tmp_path / "salida1")
    df2, g2 = _correr(cargada.plan, tmp_path / "salida2")
    assert df1.equals(df2)
    assert len(g1.parts) > 1  # se dividió por región
    assert _contenido(g1) == _contenido(g2)
    assert g1.rows_total == len(df1)


def test_rutas_relativas_si_estan_junto_a_la_receta(carpeta):
    plan = _plan(carpeta, carpeta / "salida")
    receta = carpeta / "mes.json"
    save_recipe(plan, receta)
    texto = receta.read_text(encoding="utf-8")
    datos = json.loads(texto)["plan"]
    assert datos["principal"]["path"] == "clientes.xlsx"
    assert datos["secondaries"][1]["path"] == "cobranzas.xlsx"
    assert datos["output"]["path"] == "salida/resultado.xlsx"
    assert "\\" not in texto
    assert json.loads(texto)["version"] == 1
    assert "Región" in texto  # ensure_ascii=False
    assert "\n  " in texto  # con indentación


def test_rutas_absolutas_si_estan_en_otra_carpeta(carpeta, tmp_path):
    plan = _plan(carpeta, tmp_path / "salida")
    receta_dir = tmp_path / "recetas"
    receta_dir.mkdir()
    receta = receta_dir / "mes.json"
    save_recipe(plan, receta)
    datos = json.loads(receta.read_text(encoding="utf-8"))["plan"]
    assert datos["principal"]["path"] == (carpeta / "clientes.xlsx").as_posix()
    cargada = load_recipe(receta)
    assert cargada.plan == plan
    assert cargada.missing == []


def test_mover_la_carpeta_entera_sigue_funcionando(carpeta, tmp_path):
    receta = carpeta / "mes.json"
    save_recipe(_plan(carpeta, carpeta / "salida"), receta)
    nueva = tmp_path / "movida"
    shutil.move(str(carpeta), str(nueva))
    cargada = load_recipe(nueva / "mes.json")
    assert cargada.missing == []
    assert cargada.plan.principal.path == nueva / "clientes.xlsx"
    assert cargada.plan.output.path == nueva / "salida" / "resultado.xlsx"
    assert cargada.plan == _plan(nueva, nueva / "salida")


def test_archivo_faltante_carga_igual(carpeta):
    receta = carpeta / "mes.json"
    save_recipe(_plan(carpeta, carpeta / "salida"), receta)
    (carpeta / "cobranzas.xlsx").unlink()
    cargada = load_recipe(receta)
    assert cargada.missing == ["cobranzas"]
    assert any("cobranzas.xlsx" in w for w in cargada.warnings)
    assert cargada.plan.secondaries[1].summary is not None


@pytest.mark.parametrize(
    "ruta", ["C:\\Datos\\ventas.xlsx", "C:/Datos/ventas.xlsx", "\\\\srv\\x\\a.csv"]
)
def test_rutas_de_windows_se_leen_como_texto_y_figuran_faltantes(tmp_path, ruta):
    receta = tmp_path / "mes.json"
    plan = JoinPlan(SourceSpec("ventas", Path("x.csv"), ReadOptions(), ["A"]), KeyOptions(), [])
    save_recipe(plan, receta)
    datos = json.loads(receta.read_text(encoding="utf-8"))
    datos["plan"]["principal"]["path"] = ruta
    receta.write_text(json.dumps(datos), encoding="utf-8")
    cargada = load_recipe(receta)
    assert cargada.missing == ["ventas"]
    assert str(cargada.plan.principal.path).endswith(ruta.replace("\\", "/").split("/")[-1]) or (
        "\\" in str(cargada.plan.principal.path)
    )


def test_version_desconocida(tmp_path):
    receta = tmp_path / "r.json"
    receta.write_text(json.dumps({"version": 2, "plan": {}}), encoding="utf-8")
    with pytest.raises(ExcelToolError, match="dañada o es de una versión más nueva"):
        load_recipe(receta)


@pytest.mark.parametrize(
    "contenido", ["{ esto no es json", "", "[1, 2]", '{"plan": {}}', '{"version": 1}', "\x00\xff"]
)
def test_json_roto_o_sin_version(tmp_path, contenido):
    receta = tmp_path / "r.json"
    receta.write_text(contenido, encoding="utf-8")
    with pytest.raises(ExcelToolError, match="dañada"):
        load_recipe(receta)


def test_archivo_no_utf8(tmp_path):
    receta = tmp_path / "r.json"
    receta.write_bytes(b'{"version": 1, "x": "\xe1"}')
    with pytest.raises(ExcelToolError, match="dañada"):
        load_recipe(receta)


def test_receta_inexistente(tmp_path):
    with pytest.raises(ExcelToolError, match="No encontré"):
        load_recipe(tmp_path / "nada.json")


def _modificar(receta: Path, fn):
    datos = json.loads(receta.read_text(encoding="utf-8"))
    fn(datos["plan"])
    receta.write_text(json.dumps(datos), encoding="utf-8")


@pytest.mark.parametrize(
    "cambio",
    [
        lambda p: p["secondaries"][0].update(mode="otro"),
        lambda p: p["secondaries"][0].update(key_pairs=[["a"]]),
        lambda p: p["secondaries"][0]["summary"]["aggregations"][0].update(func="mediana"),
        lambda p: p["secondaries"][1]["summary"]["group_by"][1].update(period="semana"),
        lambda p: p["output"].update(format="pdf"),
        lambda p: p["output"]["split"].update(rows_per_part="mil"),
        lambda p: p["key_options"].update(trim="si"),
        lambda p: p.pop("principal"),
        lambda p: p["principal"].update(path=5),
        lambda p: p.update(secondaries={}),
    ],
)
def test_datos_invalidos(carpeta, tmp_path, cambio):
    receta = carpeta / "mes.json"
    save_recipe(_plan(carpeta, carpeta / "salida"), receta)
    _modificar(receta, cambio)
    with pytest.raises(ExcelToolError, match="dañada"):
        load_recipe(receta)


def test_campos_extra_se_ignoran_con_aviso(carpeta, caplog):
    receta = carpeta / "mes.json"
    plan = _plan(carpeta, carpeta / "salida")
    save_recipe(plan, receta)

    def agregar(p):
        p["campo_nuevo"] = 1
        p["secondaries"][0]["otra_cosa"] = {"a": 1}
        p["output"]["split"]["futuro"] = True

    _modificar(receta, agregar)
    with caplog.at_level("WARNING"):
        cargada = load_recipe(receta)
    assert cargada.plan == plan
    assert cargada.warnings
    assert "campo_nuevo" in caplog.text and "otra_cosa" in caplog.text


def test_plan_minimo_sin_opcionales(tmp_path):
    plan = JoinPlan(
        SourceSpec("a", tmp_path / "a.csv", ReadOptions(), ["X", "Y"]),
        KeyOptions(),
        [SecondarySpec("b", tmp_path / "b.csv", ReadOptions(), [], [("X", "X"), ("Y", "Z")])],
    )
    receta = tmp_path / "r.json"
    save_recipe(plan, receta)
    cargada = load_recipe(receta)
    assert cargada.plan == plan
    assert cargada.plan.secondaries[0].key_pairs == [("X", "X"), ("Y", "Z")]
    assert cargada.missing == ["a", "b"]


def test_sobrescribe_receta_existente(tmp_path):
    receta = tmp_path / "r.json"
    plan = JoinPlan(SourceSpec("a", tmp_path / "a.csv", ReadOptions(), []), KeyOptions(), [])
    save_recipe(plan, receta)
    save_recipe(replace(plan, output_column_order=["Z"]), receta)
    assert load_recipe(receta).plan.output_column_order == ["Z"]
    assert [p.name for p in tmp_path.iterdir()] == ["r.json"]


def test_guardar_en_carpeta_inexistente(tmp_path):
    plan = JoinPlan(SourceSpec("a", tmp_path / "a.csv", ReadOptions(), []), KeyOptions(), [])
    with pytest.raises(ExcelToolError, match="No se pudo guardar"):
        save_recipe(plan, tmp_path / "no_existe" / "r.json")

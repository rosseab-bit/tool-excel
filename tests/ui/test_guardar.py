"""Paso Guardar (F3-10). El core de la Fase 3 se reemplaza por dobles del contrato."""

from __future__ import annotations

import os
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

from excel_tool import platform_utils
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.preferences import load_preferences
from excel_tool.ui import guardar

pytestmark = pytest.mark.pantalla


# --- dobles con los mismos campos que docs/contratos.md § Fase 3 ---
@dataclass(frozen=True)
class SplitSpec:
    mode: str
    rows_per_part: int | None = None
    group_column: str | None = None
    value_columns: list[str] | None = None
    name_pattern: str | None = None


@dataclass(frozen=True)
class OutputSpec:
    path: Path
    format: str
    destination: str
    split: SplitSpec | None = None
    write_index: bool = True
    sheet_name: str = "Datos"
    column_types: dict | None = None


@dataclass(frozen=True)
class PartPlan:
    name: str
    rows: int
    value: tuple[str, ...] | None
    first_row: int
    last_row: int


@dataclass(frozen=True)
class PartResult:
    path: Path
    sheet: str | None
    rows: int
    value: tuple[str, ...] | None
    first_row: int
    last_row: int


@dataclass(frozen=True)
class SaveResult:
    files: list[Path]
    folder: Path | None
    parts: list[PartResult]
    rows_total: int
    warnings: list[str] = field(default_factory=list)
    unconverted: dict[str, int] = field(default_factory=dict)


@dataclass(frozen=True)
class SizeCheck:
    level: str
    reasons: list[str]
    suggested_parts: int
    suggested_rows_per_part: int
    suggested_columns: list


@dataclass(frozen=True)
class ColumnSuggestion:
    column: str
    parts: int
    largest_part_rows: int
    fits: bool


class CoreFalso:
    def __init__(self, nivel="ok", partes=2):
        self.nivel = nivel
        self.n_partes = partes
        self.checks = []
        self.planes = []
        self.guardados = []
        self.comportamiento_guardar = None

    def estimate(self, df):
        return SimpleNamespace(rows=len(df))

    def check(self, est, destino, formato, *, limits=None, suggestions=None):
        self.checks.append((destino, formato, suggestions))
        razones = {
            "ok": [],
            "recomendado": ["Conviene dividirlo.", "Tiene 400.000 filas."],
            "obligatorio": [f"Demasiado grande para {destino}.", "Tiene 2.300.000 filas."],
        }[self.nivel]
        return SizeCheck(self.nivel, razones, 8, 290_000, suggestions or [])

    def suggest_split_columns(self, df, destino, formato, *, progress=None, cancel=None, **kw):
        return [ColumnSuggestion("Region", 5, 420_000, True), ColumnSuggestion("Mes", 12, 9, False)]

    def plan_split(self, df, spec, *, base_name, max_rows_per_part=None):
        self.planes.append((spec, base_name, max_rows_per_part))
        return [
            PartPlan(f"{base_name}_parte_{i + 1:02d}", 2, None, 2 * i + 1, 2 * i + 2)
            for i in range(self.n_partes)
        ]

    def save_result(self, df, spec, *, progress=None, cancel=None):
        self.guardados.append(spec)
        if self.comportamiento_guardar is not None:
            return self.comportamiento_guardar(df, spec, progress, cancel)
        carpeta = spec.path.parent / spec.path.stem if spec.split else None
        base = carpeta or spec.path.parent
        partes = [
            PartResult(base / f"p{i}{spec.path.suffix}", None, len(df), None, 1, len(df))
            for i in range(self.n_partes if spec.split else 1)
        ]
        return SaveResult([p.path for p in partes], carpeta, partes, len(df))

    def namespace(self):
        return SimpleNamespace(
            estimate=self.estimate,
            check=self.check,
            suggest_split_columns=self.suggest_split_columns,
            plan_split=self.plan_split,
            SplitSpec=SplitSpec,
            OutputSpec=OutputSpec,
            save_result=self.save_result,
        )


def _esperar(raiz, condicion, segundos=10):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _listo(v):
    return v._check is not None and (v._leer_division().spec is None or v._partes is not None)


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


@pytest.fixture
def df():
    return pd.DataFrame({"Cliente": list("abcd"), "Region": ["N", "S", "N", "S"]})


def _abrir(raiz, monkeypatch, df, nivel="ok", **kw):
    core = CoreFalso(nivel)
    monkeypatch.setattr(guardar, "_core", core.namespace)
    v = guardar.abrir_guardar(raiz, df, nombre_sugerido="resultado_ventas", **kw)
    _esperar(raiz, lambda: _listo(v))
    return v, core


def _elegir_ruta(monkeypatch, ruta):
    llamadas = []

    def falso(**kwargs):
        llamadas.append(kwargs)
        return str(ruta)

    monkeypatch.setattr(guardar.filedialog, "asksaveasfilename", falso)
    return llamadas


def _guardar_y_esperar(raiz, v):
    v.btn_guardar.invoke()
    _esperar(raiz, lambda: v._resultado is not None or v.lbl_estado.cget("text") != "Guardando…")


def test_excel_por_defecto_y_recuerda_google_sheets(raiz, monkeypatch, df):
    v, _ = _abrir(raiz, monkeypatch, df)
    assert v.v_destino.get() == "excel"
    v.rb_destino["google_sheets"].invoke()
    _esperar(raiz, lambda: _listo(v))
    assert load_preferences().last_destination == "google_sheets"
    v.cerrar()
    v2, _ = _abrir(raiz, monkeypatch, df)
    assert v2.v_destino.get() == "google_sheets"
    assert v2.v_formato.get() == "csv_google"


def test_nivel_ok_guarda_en_un_solo_archivo(raiz, monkeypatch, df, tmp_path):
    v, core = _abrir(raiz, monkeypatch, df, "ok")
    assert v.v_modo.get() == "none"
    assert "entra sin problemas" in v.lbl_principal.cget("text")
    llamadas = _elegir_ruta(monkeypatch, tmp_path / "salida" / "res.xlsx")
    _guardar_y_esperar(raiz, v)
    spec = core.guardados[0]
    assert spec.split is None
    assert (spec.path, spec.format, spec.destination) == (
        tmp_path / "salida" / "res.xlsx",
        "xlsx",
        "excel",
    )
    assert llamadas[0]["initialfile"] == "resultado_ventas.xlsx"
    assert load_preferences().last_folder == tmp_path / "salida"


def test_recomendado_preselecciona_dividir_pero_deja_un_archivo(raiz, monkeypatch, df, tmp_path):
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    assert v.v_modo.get() == "rows"
    assert v.v_filas.get() == "290.000"
    assert v.lbl_principal.cget("text") == "Conviene dividirlo."
    assert v.puede_guardar()
    v.rb_uno.invoke()
    _esperar(raiz, lambda: _listo(v))
    assert v.puede_guardar()
    _elegir_ruta(monkeypatch, tmp_path / "x.xlsx")
    _guardar_y_esperar(raiz, v)
    assert core.guardados[0].split is None


def test_recomendado_con_claves_sugeridas_preselecciona_sin_cortar(raiz, monkeypatch, df):
    v, _ = _abrir(raiz, monkeypatch, df, "recomendado", claves_sugeridas=["Cliente"])
    assert v.v_modo.get() == "keep"
    assert v.v_col_grupo.get() == "Cliente"


def test_obligatorio_no_deja_un_solo_archivo(raiz, monkeypatch, df):
    v, _ = _abrir(raiz, monkeypatch, df, "obligatorio")
    assert "disabled" in v.rb_uno.state()
    assert v.v_modo.get() != "none"
    assert "NO ENTRA" in v.lbl_nivel.cget("text")
    assert "no se puede guardar en un solo archivo" in v.lbl_detalle.cget("text").lower()
    assert v.puede_guardar()
    v.v_modo.set("none")
    assert not v.puede_guardar()


def test_varias_hojas_no_aparece_con_google_sheets(raiz, monkeypatch, df):
    v, _ = _abrir(raiz, monkeypatch, df)
    assert v.rb_hojas.winfo_manager() != ""
    v.rb_destino["google_sheets"].invoke()
    _esperar(raiz, lambda: _listo(v))
    assert v.rb_hojas.winfo_manager() == ""
    v.rb_destino["excel"].invoke()
    _esperar(raiz, lambda: _listo(v))
    assert v.rb_hojas.winfo_manager() != ""
    # con CSV no se puede repartir en hojas
    v.rb_formato["csv_excel"].invoke()
    assert "disabled" in v.rb_hojas.state()


def test_formatos_segun_el_destino(raiz, monkeypatch, df):
    v, _ = _abrir(raiz, monkeypatch, df)
    assert set(v.rb_formato) == {"xlsx", "csv_excel"}
    v.rb_destino["google_sheets"].invoke()
    assert set(v.rb_formato) == {"xlsx", "csv_google"}
    v.rb_destino["solo_guardar"].invoke()
    assert set(v.rb_formato) == {"csv_standard"}
    textos = [rb.cget("text") for rb in v.rb_formato.values()]
    assert not any("encoding" in t.lower() for t in textos)


def test_vista_previa_muestra_las_partes(raiz, monkeypatch, df):
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    filas = [v.tabla.item(i, "values") for i in v.tabla.get_children()]
    assert filas[0] == ("resultado_ventas_parte_01", "2", "1", "2")
    assert len(filas) == 2
    assert "2 archivos" in v.lbl_vista.cget("text")
    spec, base, _ = core.planes[-1]
    assert spec.mode == "rows" and spec.rows_per_part == 290_000 and base == "resultado_ventas"


def test_vista_previa_limita_a_200(raiz, monkeypatch, df):
    core = CoreFalso("recomendado", partes=250)
    monkeypatch.setattr(guardar, "_core", core.namespace)
    v = guardar.abrir_guardar(raiz, df, nombre_sugerido="r")
    _esperar(raiz, lambda: _listo(v))
    assert len(v.tabla.get_children()) == 200
    assert "250" in v.lbl_vista.cget("text")


def test_sugerencias_de_columnas_para_dividir(raiz, monkeypatch, df):
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    textos = list(v.cmb_valor.cget("values"))
    assert textos[0] == "Region: 5 archivos, el más grande de 420.000 filas ✅"
    assert textos[1].endswith("⚠️")
    assert core.checks[-1][2] is not None  # el semáforo recibió las sugerencias


def test_tabla_grande_busca_columnas_solo_a_pedido(raiz, monkeypatch, df):
    monkeypatch.setattr(guardar, "AUTO_FILAS", 2)
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    assert list(v.cmb_valor.cget("values")) == []
    v.btn_buscar.invoke()
    _esperar(raiz, lambda: _listo(v) and v.cmb_valor.cget("values"))
    assert len(v.cmb_valor.cget("values")) == 2


@pytest.mark.parametrize(
    ("modo", "esperado"),
    [
        ("rows", SplitSpec(mode="rows", rows_per_part=1000)),
        ("keep", SplitSpec(mode="rows_keep_groups", rows_per_part=1000, group_column="Cliente")),
        ("value", SplitSpec(mode="by_value", value_columns=["Region"], rows_per_part=1000)),
        ("sheets", SplitSpec(mode="sheets", rows_per_part=1000)),
    ],
)
def test_guardar_arma_el_output_spec_de_cada_modo(raiz, monkeypatch, df, tmp_path, modo, esperado):
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    v.v_filas.set("1.000")
    v.v_col_grupo.set("Cliente")
    v.v_modo.set(modo)
    v._etapa_b()
    _esperar(raiz, lambda: _listo(v))
    _elegir_ruta(monkeypatch, tmp_path / "res.xlsx")
    _guardar_y_esperar(raiz, v)
    spec = core.guardados[0]
    assert spec.split == esperado
    assert spec.format == "xlsx" and spec.destination == "excel"


def test_por_valor_sin_subdividir(raiz, monkeypatch, df):
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    v.v_modo.set("value")
    v.v_subdividir.set(False)
    v._etapa_b()
    _esperar(raiz, lambda: _listo(v))
    spec, _, tope = core.planes[-1]
    assert spec.rows_per_part is None and tope is None


def test_filas_invalidas_no_deja_guardar(raiz, monkeypatch, df):
    v, _ = _abrir(raiz, monkeypatch, df, "recomendado")
    v.v_filas.set("abc")
    v._etapa_b()
    assert not v.puede_guardar()
    assert "cuántas filas" in v.lbl_vista.cget("text")


def test_resultado_y_abrir_carpeta(raiz, monkeypatch, df, tmp_path):
    v, core = _abrir(raiz, monkeypatch, df, "recomendado")
    terminados = []
    v.on_terminado = terminados.append
    abiertas = []
    monkeypatch.setattr(platform_utils, "open_folder", abiertas.append)
    core.comportamiento_guardar = lambda d, s, p, c: SaveResult(
        [s.path.parent / "a.xlsx", s.path.parent / "b.xlsx"],
        s.path.parent / "res",
        [
            PartResult(s.path.parent / "a.xlsx", None, 2300000, None, 1, 2300000),
            PartResult(s.path.parent / "b.xlsx", None, 0, None, 1, 0),
        ],
        2300000,
        warnings=["La parte 3 pesa 120 MB."],
        unconverted={"Importe": 3},
    )
    _elegir_ruta(monkeypatch, tmp_path / "res.xlsx")
    _guardar_y_esperar(raiz, v)
    texto = v.lbl_resultado.cget("text")
    assert (
        texto
        == "Se guardaron 2 archivos con 2.300.000 filas en total (coincide con el resultado ✅)"
    )
    assert "La parte 3 pesa 120 MB." in v.lbl_avisos.cget("text")
    assert "Importe" in v.lbl_avisos.cget("text")
    assert len(terminados) == 1
    v.btn_carpeta.invoke()
    assert abiertas == [tmp_path / "res"]


def test_cancelar_no_deja_la_ventana_colgada(raiz, monkeypatch, df, tmp_path):
    v, core = _abrir(raiz, monkeypatch, df, "ok")
    liberar = threading.Event()

    def colgado(d, s, progress, cancel):
        liberar.wait(5)
        cancel.check()
        return SaveResult([], None, [], 0)

    core.comportamiento_guardar = colgado
    _elegir_ruta(monkeypatch, tmp_path / "res.xlsx")
    v.btn_guardar.invoke()

    def boton_cancelar():
        for hija in v.win.winfo_children():
            if isinstance(hija, tk.Toplevel):
                for nieta in hija.winfo_children():
                    for b in nieta.winfo_children():
                        if b.winfo_class() == "TButton" and b.cget("text") == "Cancelar":
                            return b
        return None

    _esperar(raiz, lambda: boton_cancelar() is not None)
    boton_cancelar().invoke()
    liberar.set()
    raiz.update()
    assert "Se canceló" in v.lbl_estado.cget("text")
    assert v.puede_guardar()
    assert v._resultado is None


def test_error_del_core_se_muestra_amigable(raiz, monkeypatch, df, tmp_path):
    v, core = _abrir(raiz, monkeypatch, df, "ok")

    def falla(d, s, progress, cancel):
        raise ExcelToolError("Cerrá el archivo en Excel y volvé a intentar.", "detalle técnico")

    core.comportamiento_guardar = falla
    avisos = []
    monkeypatch.setattr(
        guardar.messagebox, "showerror", lambda titulo, mensaje, **kw: avisos.append(mensaje)
    )
    _elegir_ruta(monkeypatch, tmp_path / "res.xlsx")
    _guardar_y_esperar(raiz, v)
    _esperar(raiz, lambda: avisos)
    assert avisos == ["Cerrá el archivo en Excel y volvé a intentar."]
    assert v.puede_guardar()


def test_error_inesperado_no_muestra_traceback(raiz, monkeypatch, df, tmp_path):
    v, core = _abrir(raiz, monkeypatch, df, "ok")
    core.comportamiento_guardar = lambda *a: (_ for _ in ()).throw(RuntimeError("boom"))
    avisos = []
    monkeypatch.setattr(
        guardar.messagebox, "showerror", lambda titulo, mensaje, **kw: avisos.append(mensaje)
    )
    _elegir_ruta(monkeypatch, tmp_path / "res.xlsx")
    v.btn_guardar.invoke()
    _esperar(raiz, lambda: avisos)
    assert "boom" not in avisos[0]


def test_sin_elegir_ruta_no_guarda(raiz, monkeypatch, df):
    v, core = _abrir(raiz, monkeypatch, df, "ok")
    monkeypatch.setattr(guardar.filedialog, "asksaveasfilename", lambda **kw: "")
    v.btn_guardar.invoke()
    assert core.guardados == []


def test_core_ausente_avisa_con_mensaje(raiz, monkeypatch, df):
    def no_hay():
        raise ExcelToolError(guardar.TEXTO_NO_DISPONIBLE)

    monkeypatch.setattr(guardar, "_core", no_hay)
    avisos = []
    monkeypatch.setattr(guardar.messagebox, "showinfo", lambda t, m, **kw: avisos.append(m))
    assert guardar.abrir_guardar(raiz, df, nombre_sugerido="r") is None
    assert avisos == [guardar.TEXTO_NO_DISPONIBLE]

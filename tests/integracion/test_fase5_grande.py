"""Escenario grande (200.000 clientes x 2,5 M de ventas, en detalle) manejando el asistente real.

Tarda varios minutos. Se saltea si no están los archivos grandes (ver `fixtures_grandes`).
"""

from __future__ import annotations

import pytest

from .asistente_util import Conductor
from .escenario_grande_ui import cargar_esperado, recorrer

pytestmark = pytest.mark.pantalla


def test_escenario_grande_desde_la_ui(
    entorno_ui, monkeypatch, fixtures_grandes, fixtures_esperados, tmp_path
):
    c = Conductor(entorno_ui, monkeypatch)
    hechos = recorrer(
        c, fixtures_grandes, tmp_path, cargar_esperado(fixtures_esperados), log=lambda *_: None
    )
    # Google Sheets: cada parte entra (celdas y peso) aunque se haya ofrecido otra cantidad
    for archivo_mb in hechos["sheets"]["mb_por_archivo"]:
        assert archivo_mb < 100
    for destino in ("excel", "sheets"):
        h = hechos[destino]
        # la cantidad que se sugiere es la que se genera (o una más) y ninguna parte queda minúscula
        assert h["partes_en_la_vista_previa"] in (
            h["partes_sugeridas"],
            h["partes_sugeridas"] + 1,
        ), (destino, h["partes_sugeridas"], h["partes_filas"])
        if destino == "excel":  # en Sheets se llenan las partes al tope y la última es lo que sobra
            minimo = max(h["partes_filas"]) // 10
            assert min(h["partes_filas"]) > minimo, (destino, h["partes_filas"])
    # F5-A: la primera vez que se abre Guardar para Excel, "sin cortar un cliente" no se desborda
    assert hechos["excel"]["partes_en_la_vista_previa"] - hechos["excel"]["partes_sugeridas"] <= 1
    # Al cambiar de Excel a Google Sheets se vuelve a la división sugerida
    cambio = hechos["sheets_heredada"]["al_cambiar"]
    assert cambio["partes"] in (cambio["partes_sugeridas"], cambio["partes_sugeridas"] + 1), cambio
    assert cambio["filas"] == cambio["filas_sugeridas"] or cambio["partes"] <= 6, cambio

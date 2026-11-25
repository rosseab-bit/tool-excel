"""Perfiles falsos con los mismos campos del contrato (core/profile.py), para probar la UI."""

from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

from excel_tool.core.io import CsvOptions, SheetSummary


def columna(nombre, posicion, tipo="texto", **extra):
    datos = {
        "name": nombre,
        "position": posicion,
        "letter": "ABCDEFGHIJ"[posicion],
        "detected_type": tipo,
        "type_breakdown": {},
        "non_empty": 1000,
        "empty": 0,
        "distinct": 5,
        "is_unique": False,
        "few_values": False,
        "summable": False,
        "top_values": [],
        "sample_values": [],
        "min": None,
        "max": None,
        "sum": None,
        "mean": None,
        "min_len": None,
        "max_len": None,
        "leading_zeros": 0,
        "untrimmed": 0,
        "case_variants": 0,
        "not_converted": 0,
        "notes": [],
    }
    datos.update(extra)
    return SimpleNamespace(**datos)


def columnas_falsas():
    return [
        columna(
            "ID cliente",
            0,
            "entero",
            non_empty=10000,
            distinct=10000,
            is_unique=True,
            summable=True,
            min="1001",
            max="11000",
            sum=60005000.0,
            mean=6000.5,
            sample_values=["1001", "1002"],
        ),
        columna(
            "Importe",
            1,
            "mixto",
            type_breakdown={"decimal": 995, "texto": 5},
            summable=True,
            not_converted=5,
            sum=1234567.891,
            mean=1234.5,
            top_values=[("10,5", 40), ("", 3)],
            notes=["Hay 5 valores que no son números."],
        ),
        columna(
            "Región",
            2,
            "texto",
            few_values=True,
            min_len=3,
            max_len=5,
            case_variants=2,
            untrimmed=1,
            leading_zeros=4,
            top_values=[("Norte", 500), ("NORTE", 2)],
            sample_values=["Norte", "Sur"],
        ),
    ]


def perfil_falso(path=Path("/datos/ventas.csv"), hojas=None, hoja=None, csv=True):
    cols = columnas_falsas()
    return SimpleNamespace(
        file=SimpleNamespace(
            name=Path(path).name,
            path=Path(path),
            size=12_900_000,
            modified=datetime(2026, 3, 5, 14, 30),
            format="csv" if csv else "xlsx",
        ),
        sheets=hojas or [],
        sheet=hoja,
        csv_options=CsvOptions(";", "utf-8-sig") if csv else None,
        header_row=0,
        rows=1234,
        columns=len(cols),
        empty_rows=2,
        duplicate_rows=7,
        columns_profile=cols,
        warnings=["La columna Importe tiene 5 valores que no son números."],
    )


def hojas_falsas():
    return [SheetSummary("Ventas", 100, 3), SheetSummary("Clientes", 50, 4)]

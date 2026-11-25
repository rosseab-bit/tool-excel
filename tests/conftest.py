"""Fixtures compartidos de los tests."""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
SCRIPT_FIXTURES = RAIZ / "scripts" / "make_fixtures.py"
GENERADOS = RAIZ / "tests" / "fixtures" / "generados"
ESPERADOS = RAIZ / "tests" / "fixtures" / "esperados"


def _cargar_generador():
    spec = importlib.util.spec_from_file_location("make_fixtures", SCRIPT_FIXTURES)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def _al_dia(generador) -> bool:
    version = GENERADOS / generador.ARCHIVO_VERSION
    if not version.is_file():
        return False
    return version.read_text(encoding="utf-8").strip() == generador.huella_generador()


@pytest.fixture(scope="session")
def generador_fixtures():
    """El módulo scripts/make_fixtures.py ya cargado."""
    return _cargar_generador()


@pytest.fixture(scope="session")
def fixtures_generados(generador_fixtures) -> Path:
    """Carpeta con los archivos de prueba. Si falta o quedó desactualizada, la regenera
    (sin los archivos grandes)."""
    if not _al_dia(generador_fixtures):
        generador_fixtures.generar(GENERADOS, sin_grandes=True)
    return GENERADOS


@pytest.fixture(scope="session")
def fixtures_esperados() -> Path:
    """Carpeta con los resultados esperados (se suben al repo)."""
    return ESPERADOS


@pytest.fixture(scope="session")
def fixtures_grandes(fixtures_generados, generador_fixtures) -> Path:
    """Carpeta con los archivos grandes. Saltea el test si no están generados
    (se generan a mano con scripts/make_fixtures.py, sin --sin-grandes)."""
    marca = fixtures_generados / "grandes" / "LISTO"
    if (
        not marca.is_file()
        or marca.read_text(encoding="utf-8").strip() != generador_fixtures.huella_generador()
    ):
        pytest.skip(
            "Faltan los archivos grandes de prueba: correr "
            "'.venv/bin/python scripts/make_fixtures.py' (sin --sin-grandes)."
        )
    return fixtures_generados / "grandes"

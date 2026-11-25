import pytest

from excel_tool.core.errors import (
    ArchivoBloqueado,
    ArchivoIlegible,
    ArchivoNoEncontrado,
    ExcelToolError,
    FormatoNoSoportado,
)


def test_mensaje_y_detalle():
    e = ExcelToolError("Algo falló", "detalle técnico")
    assert e.mensaje == "Algo falló"
    assert e.detalle == "detalle técnico"
    assert str(e) == "Algo falló"


def test_detalle_opcional():
    assert ExcelToolError("hola").detalle is None


@pytest.mark.parametrize(
    "clase", [ArchivoNoEncontrado, ArchivoBloqueado, FormatoNoSoportado, ArchivoIlegible]
)
def test_subclases(clase):
    assert issubclass(clase, ExcelToolError)
    with pytest.raises(ExcelToolError):
        raise clase("x", "y")

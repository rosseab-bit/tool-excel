"""Los 7 pasos del asistente, en orden."""

from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.unir.pasos.paso1_principal import PasoPrincipal
from excel_tool.ui.unir.pasos.paso2_clave import PasoClave
from excel_tool.ui.unir.pasos.paso3_sumar import PasoSumar
from excel_tool.ui.unir.pasos.paso4_modo import PasoModo
from excel_tool.ui.unir.pasos.paso5_columnas import PasoColumnas
from excel_tool.ui.unir.pasos.paso6_resumen import PasoResumenFinal
from excel_tool.ui.unir.pasos.paso7_guardar import PasoGuardar

PASOS: tuple[type[Paso], ...] = (
    PasoPrincipal,
    PasoClave,
    PasoSumar,
    PasoModo,
    PasoColumnas,
    PasoResumenFinal,
    PasoGuardar,
)

__all__ = ["PASOS", "Paso"]

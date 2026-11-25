from datetime import datetime

from excel_tool.ui import formato

from .perfiles_falsos import columnas_falsas, perfil_falso


def test_numeros_argentinos():
    assert formato.formatear_decimal(1234567.891) == "1.234.567,89"
    assert formato.formatear_valor(1000.0) == "1.000"
    assert formato.formatear_valor(-1234.5) == "-1.234,50"
    assert formato.formatear_porcentaje(99.5) == "99,5 %"


def test_tamano_legible():
    assert formato.formatear_tamano(500) == "500 bytes"
    assert formato.formatear_tamano(12_900_000) == "12,3 MB"
    assert formato.formatear_tamano(1536) == "1,5 KB"
    assert formato.formatear_tamano(3 * 1024**3) == "3,0 GB"


def test_fecha_y_tipos():
    assert formato.formatear_fecha(datetime(2026, 3, 5, 9, 5)) == "05/03/2026 09:05"
    assert formato.nombre_tipo("entero") == "Número entero"
    assert formato.nombre_tipo("fecha_hora") == "Fecha y hora"
    assert formato.nombre_tipo("mixto") == "Mixto"


def test_fila_y_marcas():
    clave, mixta, region = columnas_falsas()
    assert formato.fila_columna(clave) == (
        "A (1)",
        "ID cliente",
        "Número entero",
        "10.000",
        "0",
        "10.000",
        "Posible clave · Se puede sumar",
    )
    assert formato.marcas_columna(region) == ["Pocos valores"]


def test_detalle_mixta_y_variantes():
    _, mixta, region = columnas_falsas()
    detalle = dict(formato.detalle_columna(mixta))
    assert "Número decimal: 99,5 % (995)" in detalle["Mezcla de tipos"]
    assert detalle["Suma"] == "1.234.567,89"
    assert detalle["No son del tipo detectado"] == "5 celdas"
    detalle = dict(formato.detalle_columna(region))
    assert "2 valores" in detalle["Escritos distinto"]
    assert detalle["Largo del texto"] == "de 3 a 5 caracteres"
    assert "Mínimo" not in detalle


def test_generales_csv_sin_jerga():
    secciones = dict(formato.datos_generales(perfil_falso()))
    assert ("Tamaño", "12,3 MB") in secciones["El archivo"]
    assert ("Filas repetidas", "7") in secciones["Los datos"]
    assert ("Tipo de texto", "Recomendado (con acentos y ñ)") in secciones["El texto"]

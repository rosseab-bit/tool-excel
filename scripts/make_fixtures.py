"""Genera los archivos de prueba (fixtures) de Excel Join Tool.

Uso:
    .venv/bin/python scripts/make_fixtures.py [--sin-grandes] [--salida DIR]

Por defecto escribe en tests/fixtures/generados/ (no se sube al repo). Es determinístico:
usa random.Random con semilla fija y nunca mira la hora. Es idempotente: volver a correrlo
pisa lo anterior. Nunca escribe fuera de la carpeta de salida.

Cómo agregar grupos nuevos: escribir una función ``def grupo_xxx(salida: Path) -> list[str]``
(devuelve los nombres relativos de los archivos que generó) y sumarla a ``GRUPOS``.
Los grupos con ``grande=True`` se saltean con --sin-grandes.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import random
import struct
import sys
import time
import zipfile
from collections.abc import Callable
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SALIDA_DEFECTO = RAIZ / "tests" / "fixtures" / "generados"
ESPERADOS = RAIZ / "tests" / "fixtures" / "esperados"
ARCHIVO_VERSION = "VERSION"
SEMILLA = 20240601
FECHA_FIJA = dt.datetime(2024, 6, 1, 12, 0, 0)


def huella_generador() -> str:
    """Huella del código de este script: si cambia, los fixtures quedan desactualizados."""
    return hashlib.sha256(Path(__file__).read_bytes()).hexdigest()[:16]


def ruta_segura(salida: Path, relativa: str) -> Path:
    """Arma la ruta dentro de `salida` y crea la carpeta. Falla si se sale de ahí."""
    destino = (salida / relativa).resolve()
    if not destino.is_relative_to(salida.resolve()):
        raise ValueError(f"Ruta fuera de la carpeta de salida: {relativa}")
    destino.parent.mkdir(parents=True, exist_ok=True)
    return destino


def carpeta_de_esperados(salida: Path) -> Path:
    """Con la salida por defecto, los esperados van a tests/fixtures/esperados/ (se suben al
    repo). Con otra salida, van a <salida>/esperados/: nunca se escribe fuera de la salida."""
    if salida.resolve() == SALIDA_DEFECTO.resolve():
        carpeta = ESPERADOS
    else:
        carpeta = ruta_segura(salida, "esperados/x").parent
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def escribir_json(destino: Path, datos: dict) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(
        json.dumps(datos, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n"
    )


# ---------------------------------------------------------------------------
# Datos base compartidos
# ---------------------------------------------------------------------------

ENCABEZADO_CLIENTES = ["Código", "Nombre", "Localidad", "Saldo"]
CLIENTES = [
    ("00123", "Ñandú Pérez", "Córdoba", 1234.56),
    ("00124", "José Muñoz", "Neuquén", 250.75),
    ("00125", "María Núñez", "Tucumán", 98765.43),
    ("00126", "Sebastián Peña", "Rosario", 0.5),
    ("00127", "Inés Ibáñez", "Mendoza", 15000.25),
    ("00128", "Ángel Güemes", "Salta", 310.1),
    ("00129", "Lucía Fernández", "Bahía Blanca", 4321.99),
    ("01000", "Cañuelas S.A.", "Cañuelas", 777.77),
    ("01001", "Dueñas y Cía", "Río Cuarto", 12.34),
    ("01002", "Óscar Sánchez", "Paraná", 56789.01),
    ("01003", "Úrsula Vázquez", "Maipú", 1000.5),
    ("01004", "Zoé Rodríguez", "Santa Fe", 2.25),
]


def _decimal_ar(valor: float) -> str:
    """1234.56 -> '1.234,56' (como exporta Excel en Argentina)."""
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _escribir_csv(destino: Path, filas, delimitador: str, encoding: str) -> None:
    with destino.open("w", encoding=encoding, newline="") as f:
        csv.writer(f, delimiter=delimitador, lineterminator="\r\n").writerows(filas)


# ---------------------------------------------------------------------------
# Grupo: formatos (xlsx, xls, xlsb)
# ---------------------------------------------------------------------------


def _hoja_ventas() -> list[list[str]]:
    rng = random.Random(SEMILLA + 1)
    articulos = ["Tornillo ñandú", "Cañería 1/2", "Lámpara LED", "Aceite", "Pegamento"]
    filas = [["Fecha", "Código", "Artículo", "Cantidad", "Precio", "Observación"]]
    for _ in range(25):
        dia = dt.date(2024, 1, 1) + dt.timedelta(days=rng.randint(0, 180))
        filas.append(
            [
                dia.strftime("%d/%m/%Y"),
                rng.choice(CLIENTES)[0],
                rng.choice(articulos),
                str(rng.randint(1, 50)),
                f"{rng.randint(100, 99999) / 100:.2f}".replace(".", ","),
                rng.choice(["", "Entrega rápida", "Sin observación", "Revisión pendiente"]),
            ]
        )
    return filas


HOJA_RESUMEN = [["Concepto", "Valor"], ["Año", "2024"], ["Observación", "Resumen anual ñandú"]]


def generar_xlsx_varias_hojas(destino: Path) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    wb.properties.creator = "make_fixtures"
    wb.properties.created = FECHA_FIJA
    wb.properties.modified = FECHA_FIJA
    ws = wb.active
    ws.title = "Clientes"
    ws.append(ENCABEZADO_CLIENTES)
    for cod, nombre, loc, saldo in CLIENTES:
        ws.append([cod, nombre, loc, saldo])
    ws2 = wb.create_sheet("Ventas")
    for fila in _hoja_ventas():
        ws2.append(fila)
    ws3 = wb.create_sheet("Resumen año")
    for fila in HOJA_RESUMEN:
        ws3.append(fila)
    wb.save(destino)


def generar_xls_chico(destino: Path) -> None:
    import xlwt

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Clientes")
    for c, titulo in enumerate(ENCABEZADO_CLIENTES):
        ws.write(0, c, titulo)
    for r, (cod, nombre, loc, saldo) in enumerate(CLIENTES, start=1):
        ws.write(r, 0, cod)
        ws.write(r, 1, nombre)
        ws.write(r, 2, loc)
        ws.write(r, 3, saldo)
    wb.save(str(destino))


# --- Escritor mínimo de .xlsb (BIFF12) ---


def _varint(n: int) -> bytes:
    """Entero de 7 bits por byte, little endian (ids y tamaños de registros BIFF12)."""
    salida = bytearray()
    while True:
        byte = n & 0x7F
        n >>= 7
        if n:
            salida.append(byte | 0x80)
        else:
            salida.append(byte)
            return bytes(salida)


def _reg(id_registro: int, datos: bytes = b"") -> bytes:
    return _varint(id_registro) + _varint(len(datos)) + datos


def _wide(texto: str) -> bytes:
    """XLWideString: cantidad de caracteres (4 bytes) + UTF-16 LE."""
    return struct.pack("<I", len(texto)) + texto.encode("utf-16-le")


def _color_auto() -> bytes:
    # BrtColor: flags (xColorType=automático), índice, tinte, RGBA
    return struct.pack("<BBhI", 0x00, 0, 0, 0)


def _styles_bin() -> bytes:
    fuente = struct.pack("<HHHHBBBB", 220, 0, 400, 0, 0, 2, 0, 0) + _color_auto()
    fuente += struct.pack("<B", 2) + _wide("Calibri")
    relleno = struct.pack("<I", 0) + _color_auto() + _color_auto() + struct.pack("<I", 0xFFFFFFFF)
    relleno += struct.pack("<d", 0) * 5 + struct.pack("<I", 0)
    borde = struct.pack("<B", 0) + (struct.pack("<H", 0) + _color_auto()) * 5
    xf_maestro = struct.pack("<HHHHHBBHH", 0xFFFF, 0, 0, 0, 0, 0, 0, 0, 0)
    xf_celda = struct.pack("<HHHHHBBHH", 0, 0, 0, 0, 0, 0, 0, 0, 0)
    estilo = struct.pack("<IHBB", 0, 0, 0, 255) + _wide("Normal")
    partes = [
        _reg(278),  # BrtBeginStyleSheet
        _reg(611, struct.pack("<I", 1)),
        _reg(43, fuente),
        _reg(612),
        _reg(603, struct.pack("<I", 2)),
        _reg(44, relleno),
        _reg(44, relleno),
        _reg(604),
        _reg(613, struct.pack("<I", 1)),
        _reg(45, borde),
        _reg(614),
        _reg(626, struct.pack("<I", 1)),
        _reg(47, xf_maestro),
        _reg(627),
        _reg(617, struct.pack("<I", 1)),
        _reg(47, xf_celda),
        _reg(618),
        _reg(619, struct.pack("<I", 1)),
        _reg(48, estilo),
        _reg(620),
        _reg(279),  # BrtEndStyleSheet
    ]
    return b"".join(partes)


def _workbook_bin(nombre_hoja: str) -> bytes:
    vista = struct.pack("<IIIIIII", 0, 0, 28800, 12840, 600, 0, 0) + struct.pack("<B", 0)
    return b"".join(
        [
            _reg(131),  # BrtBeginBook
            _reg(153, struct.pack("<II", 0, 0) + _wide("ThisWorkbook")),  # BrtWbProp
            _reg(135),  # BrtBeginBookViews
            _reg(158, vista),  # BrtBookView
            _reg(136),
            _reg(143),  # BrtBeginBundleShs
            _reg(156, struct.pack("<II", 0, 1) + _wide("rId1") + _wide(nombre_hoja)),
            _reg(144),
            _reg(132),  # BrtEndBook
        ]
    )


def _celda(columna: int, id_registro: int, valor: bytes) -> bytes:
    # Cell: columna (4 bytes) + estilo 0 (3 bytes) + flags (1 byte)
    return _reg(id_registro, struct.pack("<I", columna) + b"\x00\x00\x00\x00" + valor)


def _sheet_bin(filas: list[list[object]]) -> bytes:
    ancho = max(len(f) for f in filas)
    partes = [
        _reg(129),  # BrtBeginSheet
        _reg(148, struct.pack("<IIII", 0, len(filas) - 1, 0, ancho - 1)),  # BrtWsDim
        _reg(145),  # BrtBeginSheetData
    ]
    for r, fila in enumerate(filas):
        # BrtRowHdr: fila, estilo, alto, 3 bytes de flags, 0 rangos de columnas
        partes.append(
            _reg(0, struct.pack("<IIH", r, 0, 0) + b"\x00\x00\x00" + struct.pack("<I", 0))
        )
        for c, valor in enumerate(fila):
            if valor is None or valor == "":
                continue
            if isinstance(valor, str):
                partes.append(_celda(c, 6, _wide(valor)))  # BrtCellSt (texto)
            else:
                partes.append(_celda(c, 5, struct.pack("<d", float(valor))))  # BrtCellReal
    partes += [_reg(146), _reg(130)]  # BrtEndSheetData, BrtEndSheet
    return b"".join(partes)


def escribir_xlsb(destino: Path, nombre_hoja: str, filas: list[list[object]]) -> None:
    """Escribe un .xlsb de una sola hoja (texto y números)."""
    tipos = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
        'relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.'
        'macroEnabled.main"/>'
        '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.'
        'binary.macroEnabled.main"/>'
        '<Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.'
        'worksheet"/>'
        '<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>'
        "</Types>"
    )
    rel_base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel_base}/officeDocument" Target="xl/workbook.bin"/>'
        "</Relationships>"
    )
    rels_wb = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel_base}/worksheet" Target="worksheets/sheet1.bin"/>'
        f'<Relationship Id="rId2" Type="{rel_base}/styles" Target="styles.bin"/>'
        "</Relationships>"
    )
    contenido = [
        ("[Content_Types].xml", tipos.encode("utf-8")),
        ("_rels/.rels", rels.encode("utf-8")),
        ("xl/workbook.bin", _workbook_bin(nombre_hoja)),
        ("xl/_rels/workbook.bin.rels", rels_wb.encode("utf-8")),
        ("xl/worksheets/sheet1.bin", _sheet_bin(filas)),
        ("xl/styles.bin", _styles_bin()),
    ]
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, datos in contenido:
            info = zipfile.ZipInfo(nombre, date_time=(2024, 6, 1, 12, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, datos)


def generar_xlsb_chico(destino: Path) -> None:
    filas: list[list[object]] = [list(ENCABEZADO_CLIENTES)]
    filas += [[cod, nombre, loc, saldo] for cod, nombre, loc, saldo in CLIENTES]
    escribir_xlsb(destino, "Clientes", filas)


def grupo_formatos(salida: Path) -> list[str]:
    generar_xlsx_varias_hojas(ruta_segura(salida, "formatos/varias_hojas.xlsx"))
    generar_xls_chico(ruta_segura(salida, "formatos/chico.xls"))
    generar_xlsb_chico(ruta_segura(salida, "formatos/chico.xlsb"))
    return ["formatos/varias_hojas.xlsx", "formatos/chico.xls", "formatos/chico.xlsb"]


# ---------------------------------------------------------------------------
# Grupo: csv
# ---------------------------------------------------------------------------


def grupo_csv(salida: Path) -> list[str]:
    encabezado = ["Código", "Nombre", "Localidad", "Importe"]
    ar = [encabezado] + [[c, n, loc, _decimal_ar(s)] for c, n, loc, s in CLIENTES]
    punto = [encabezado] + [[c, n, loc, f"{s:.2f}"] for c, n, loc, s in CLIENTES]
    _escribir_csv(ruta_segura(salida, "csv/excel_ar.csv"), ar, ";", "cp1252")
    _escribir_csv(ruta_segura(salida, "csv/coma_utf8.csv"), punto, ",", "utf-8")
    _escribir_csv(ruta_segura(salida, "csv/con_bom.csv"), punto, ",", "utf-8-sig")
    _escribir_csv(ruta_segura(salida, "csv/tabulado.txt"), punto, "\t", "utf-8")
    return ["csv/excel_ar.csv", "csv/coma_utf8.csv", "csv/con_bom.csv", "csv/tabulado.txt"]


# ---------------------------------------------------------------------------
# Grupo: explorar (para_explorar.csv / .xlsx + perfil.json)
# ---------------------------------------------------------------------------

# Cada celda se genera con su tipo de construcción. El perfil esperado se calcula
# contando esas etiquetas, sin leer los archivos ni usar código del producto.
COLUMNAS_EXPLORAR = [
    # (nombre, tipo esperado de la columna)
    ("ID", "entero"),
    ("Importe", "decimal"),
    ("Fecha", "fecha"),
    ("Fecha y hora", "fecha_hora"),
    ("Activo", "si_no"),
    ("Descripción", "texto"),
    ("Código", "texto"),
    ("Mixta", "mixto"),
    ("Vacía", "vacia"),
    ("", "texto"),
    ("Total", "entero"),
    ("Total", "decimal"),
    ("Región", "texto"),
]
BASE_FILAS = 188
CANT_REPETIDAS = 8
CANT_VACIAS = 4

Celda = tuple[str, str]  # (valor, etiqueta de tipo; "vacio" si no hay valor)


def _filas_explorar() -> list[list[Celda]]:
    rng = random.Random(SEMILLA + 2)
    palabras = ["Ñandú", "Cañería", "Lámpara", "Aceite", "Tornillo", "Pegamento", "Camión", "Árbol"]
    detalles = ["grande", "chico", "usado", "nuevo", "económico", "reforzado"]
    observaciones = ["revisar", "urgente", "sin novedad", "pendiente de pago", "cliente nuevo"]
    regiones = ["Norte", "NORTE", "Norte ", "Sur", "Centro", "Cuyo"]

    def con_hueco(valor: str, tipo: str, prob: float = 0.06) -> Celda:
        return ("", "vacio") if rng.random() < prob else (valor, tipo)

    base: list[list[Celda]] = []
    for i in range(1, BASE_FILAS + 1):
        dia = dt.date(2024, 1, 1) + dt.timedelta(days=rng.randint(0, 365))
        hora = f"{rng.randint(0, 23):02d}:{rng.randint(0, 59):02d}"
        mixta = (
            (str(rng.randint(1, 9999)), "entero")
            if rng.random() < 0.8
            else (rng.choice(["s/d", "pendiente", "N/A"]), "texto")
        )
        base.append(
            [
                (str(i), "entero"),
                con_hueco(f"{rng.randint(100, 999999) / 100:.2f}".replace(".", ","), "decimal"),
                con_hueco(dia.strftime("%d/%m/%Y"), "fecha"),
                (f"{dia.strftime('%d/%m/%Y')} {hora}", "fecha_hora"),
                con_hueco(rng.choice(["Sí", "No"]), "si_no"),
                (f"{rng.choice(palabras)} {rng.choice(detalles)}", "texto"),
                (f"{rng.randint(1, 9999):05d}", "texto"),
                mixta,
                ("", "vacio"),
                con_hueco(rng.choice(observaciones), "texto", 0.3),
                (str(rng.randint(1, 500)), "entero"),
                (f"{rng.randint(100, 99999) / 100:.2f}".replace(".", ","), "decimal"),
                (rng.choice(regiones), "texto"),
            ]
        )
    posiciones = rng.sample(range(BASE_FILAS), CANT_REPETIDAS + CANT_VACIAS)
    repetir = set(posiciones[:CANT_REPETIDAS])
    vaciar = set(posiciones[CANT_REPETIDAS:])
    filas: list[list[Celda]] = []
    for idx, fila in enumerate(base):
        filas.append(fila)
        if idx in repetir:
            filas.append(list(fila))
        if idx in vaciar:
            filas.append([("", "vacio")] * len(COLUMNAS_EXPLORAR))
    return filas


def _perfil_explorar(filas: list[list[Celda]]) -> dict:
    total = len(filas)
    vacias = sum(1 for f in filas if all(v == "" for v, _ in f))
    vistas: set[tuple[str, ...]] = set()
    repetidas = 0
    for f in filas:
        clave = tuple(v for v, _ in f)
        if all(v == "" for v in clave):
            continue
        if clave in vistas:
            repetidas += 1
        vistas.add(clave)
    columnas = []
    for pos, (nombre, tipo) in enumerate(COLUMNAS_EXPLORAR):
        celdas = [f[pos] for f in filas]
        con_valor = [(v, t) for v, t in celdas if t != "vacio"]
        desglose: dict[str, int] = {}
        for _, t in con_valor:
            desglose[t] = desglose.get(t, 0) + 1
        letra = chr(ord("A") + pos)
        columnas.append(
            {
                "nombre": nombre,
                "posicion": pos + 1,
                "letra": letra,
                "tipo": tipo,
                "vacios": total - len(con_valor),
                "distintos": len({v for v, _ in con_valor}),
                "desglose_tipos": dict(sorted(desglose.items())) if tipo == "mixto" else None,
                "ceros_a_la_izquierda": sum(1 for v, _ in con_valor if v.startswith("0")),
                "con_espacios_en_los_bordes": sum(1 for v, _ in con_valor if v != v.strip()),
            }
        )
    return {
        "_nota": (
            "Los nombres de campos siguen el borrador de PLAN seccion 2 y se ajustan en la "
            "Fase 2. Calculado desde la definicion de los datos del generador, no leyendo "
            "el archivo. 'filas' cuenta las filas de datos (sin encabezado). "
            "'filas_repetidas' son copias exactas de una fila anterior no vacia; las filas "
            "totalmente vacias se cuentan aparte en 'filas_vacias'. 'distintos' cuenta "
            "valores no vacios tal cual estan escritos (sin recortar espacios ni unificar "
            "mayusculas). 'vacios' incluye las filas totalmente vacias."
        ),
        "archivos": ["explorar/para_explorar.csv", "explorar/para_explorar.xlsx"],
        "filas": total,
        "columnas": len(COLUMNAS_EXPLORAR),
        "filas_vacias": vacias,
        "filas_repetidas": repetidas,
        "columnas_detalle": columnas,
    }


def grupo_explorar(salida: Path) -> list[str]:
    filas = _filas_explorar()
    encabezado = [n for n, _ in COLUMNAS_EXPLORAR]
    valores = [[v for v, _ in f] for f in filas]
    _escribir_csv(
        ruta_segura(salida, "explorar/para_explorar.csv"), [encabezado] + valores, ";", "utf-8-sig"
    )
    from openpyxl import Workbook

    wb = Workbook()
    wb.properties.creator = "make_fixtures"
    wb.properties.created = FECHA_FIJA
    wb.properties.modified = FECHA_FIJA
    ws = wb.active
    ws.title = "Datos"
    ws.append([n if n != "" else None for n in encabezado])
    for fila in valores:
        ws.append([v if v != "" else None for v in fila])
    wb.save(ruta_segura(salida, "explorar/para_explorar.xlsx"))

    carpeta_esperados = carpeta_de_esperados(salida)
    perfil = _perfil_explorar(filas)
    (carpeta_esperados / "perfil.json").write_text(
        json.dumps(perfil, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n"
    )
    return ["explorar/para_explorar.csv", "explorar/para_explorar.xlsx"]


# ---------------------------------------------------------------------------
# Grupo: escenario_3 (clientes.xlsx + ventas.csv + cobranzas.xlsx)
# ---------------------------------------------------------------------------

REGIONES = ["Norte", "Sur", "Centro", "Cuyo", "Litoral"]
VENDEDORES = ["Rocío Álvarez", "Martín Gómez", "Nicolás Ibáñez", "Valeria Muñoz"]
NOMBRES = ["Ana", "Beto", "Cecilia", "Dámaso", "Estela", "Fabián", "Griselda", "Héctor", "Irene",
           "Joaquín"]  # fmt: skip
APELLIDOS = ["Núñez", "Peña", "Muñoz", "Ibáñez", "Sánchez", "Cañete"]
ARTICULOS = ["Tornillo", "Cañería", "Lámpara", "Aceite", "Pegamento", "Candado"]
MEDIOS = ["Transferencia", "Efectivo", "Cheque"]
SIN_VENTAS = {5, 17, 28}
SIN_COBRANZAS = {3, 17, 22}
HUERFANAS_VENTAS = ["99", "98"]
HUERFANA_COBRANZAS = 77
INVALIDOS_VENTAS = ["n/a", "", "12,34,5", "abc"]


def _cents(c: int) -> str:
    """Centavos -> '1234.56' sin pasar por float."""
    return f"{c // 100}.{c % 100:02d}"


def _fecha_ar(d: dt.date) -> str:
    return d.strftime("%d/%m/%Y")


def _clave(k: str) -> tuple[int, str]:
    return (len(k), k)


def datos_escenario_3() -> dict:
    """Define el escenario completo. Todo lo esperado se calcula desde acá."""
    rng = random.Random(SEMILLA + 3)
    clientes = [
        {
            "id": str(i),
            "nombre": f"{NOMBRES[(i - 1) % 10]} {APELLIDOS[(i - 1) % 6]}",
            "region": rng.choice(REGIONES),
            "vendedor": rng.choice(VENDEDORES),
        }
        for i in range(1, 31)
    ]

    def fecha() -> dt.date:
        return dt.date(2024, 1, 1) + dt.timedelta(days=rng.randint(0, 180))

    ventas: list[dict] = []
    for c in clientes:
        if int(c["id"]) in SIN_VENTAS:
            continue
        for _ in range(rng.randint(1, 5)):
            ventas.append({"id": c["id"], "centavos": rng.randint(1000, 900000)})
    for h in HUERFANAS_VENTAS:
        ventas.append({"id": h, "centavos": rng.randint(1000, 900000)})
    for v in ventas:
        if v["centavos"] % 100 == 0:
            v["centavos"] += 7
        v["fecha"] = fecha()
        v["articulo"] = rng.choice(ARTICULOS)
        v["cantidad"] = rng.randint(1, 20)
    rng.shuffle(ventas)
    propios = [n for n, v in enumerate(ventas) if v["id"] not in HUERFANAS_VENTAS]
    for pos, texto in zip(sorted(rng.sample(propios, 4)), INVALIDOS_VENTAS, strict=True):
        ventas[pos]["centavos"] = None
        ventas[pos]["importe_texto"] = texto
    for v in ventas:
        if v["centavos"] is not None:
            v["importe_texto"] = _decimal_ar(v["centavos"] / 100)

    cobranzas: list[dict] = []
    for c in clientes:
        if int(c["id"]) in SIN_COBRANZAS:
            continue
        for _ in range(rng.randint(1, 4)):
            cobranzas.append({"id": c["id"]})
    cobranzas += [{"id": str(HUERFANA_COBRANZAS)}, {"id": str(HUERFANA_COBRANZAS)}]
    for cb in cobranzas:
        cb["clave_texto"] = f" {cb['id']} " if rng.random() < 0.1 else f"{int(cb['id']):05d}"
        cb["fecha"] = fecha()
        cb["medio"] = rng.choice(MEDIOS)
        centavos = rng.randint(1000, 500000)
        cb["centavos"] = centavos + 7 if centavos % 100 == 0 else centavos
    rng.shuffle(cobranzas)
    return {"clientes": clientes, "ventas": ventas, "cobranzas": cobranzas}


def grupo_escenario_3(salida: Path) -> list[str]:
    from openpyxl import Workbook

    d = datos_escenario_3()

    def libro() -> Workbook:
        wb = Workbook()
        wb.properties.creator = "make_fixtures"
        wb.properties.created = FECHA_FIJA
        wb.properties.modified = FECHA_FIJA
        return wb

    wb = libro()
    ws = wb.active
    ws.title = "Clientes"
    ws.append(["ID cliente", "Nombre", "Región", "Vendedor"])
    for c in d["clientes"]:
        ws.append([c["id"], c["nombre"], c["region"], c["vendedor"]])
    wb.save(ruta_segura(salida, "escenario_3/clientes.xlsx"))

    filas = [["ID cliente", "Fecha", "Artículo", "Cantidad", "Importe"]]
    for v in d["ventas"]:
        filas.append(
            [v["id"], _fecha_ar(v["fecha"]), v["articulo"], str(v["cantidad"]), v["importe_texto"]]
        )
    _escribir_csv(ruta_segura(salida, "escenario_3/ventas.csv"), filas, ";", "cp1252")

    wb = libro()
    ws = wb.active
    ws.title = "Cobranzas"
    ws.append(["Nro Cliente", "Fecha", "Medio de pago", "Monto"])
    for cb in d["cobranzas"]:
        ws.append([cb["clave_texto"], _fecha_ar(cb["fecha"]), cb["medio"], cb["centavos"] / 100])
    wb.save(ruta_segura(salida, "escenario_3/cobranzas.xlsx"))

    esp = carpeta_de_esperados(salida) / "escenario_3"
    for nombre, contenido in esperados_escenario_3(d).items():
        escribir_json(esp / f"{nombre}.json", contenido)
    return [f"escenario_3/{n}" for n in ("clientes.xlsx", "ventas.csv", "cobranzas.xlsx")]


def _normalizar_clave(k: str, variante: str) -> str:
    k = k.strip()
    return k.lstrip("0") or "0" if variante == "sin_ceros_a_la_izquierda" else k


def _diagnostico(principal: list[str], secundaria: list[str]) -> dict:
    cuenta: dict[str, int] = {}
    for k in secundaria:
        cuenta[k] = cuenta.get(k, 0) + 1
    con = [k for k in principal if cuenta.get(k)]
    sin = [k for k in principal if not cuenta.get(k)]
    return {
        "claves_principal": len(principal),
        "claves_con_coincidencia": len(con),
        "claves_sin_coincidencia": sin,
        "claves_del_secundario_sin_pareja": sorted(
            (k for k in cuenta if k not in principal), key=_clave
        ),
        "claves_repetidas_en_principal": sorted(
            {k for k in principal if principal.count(k) > 1}, key=_clave
        ),
        "claves_con_varias_filas_en_secundario": sum(1 for k in principal if cuenta.get(k, 0) > 1),
        "filas_secundario": len(secundaria),
        "filas_resultado_detalle": sum(max(1, cuenta.get(k, 0)) for k in principal),
    }


def esperados_escenario_3(d: dict) -> dict[str, dict]:
    ids = [c["id"] for c in d["clientes"]]
    ventas, cobranzas = d["ventas"], d["cobranzas"]

    # --- resumen de ventas
    por_cliente: dict[str, dict] = {}
    for v in ventas:
        r = por_cliente.setdefault(
            v["id"], {"filas": 0, "centavos": 0, "invalidos": 0, "arts": set(), "ultima": None}
        )
        r["filas"] += 1
        r["centavos"] += v["centavos"] or 0
        r["invalidos"] += v["centavos"] is None
        r["arts"].add(v["articulo"])
        r["ultima"] = max(r["ultima"] or v["fecha"], v["fecha"])
    resumen_ventas = {
        "_nota": (
            "Calculado desde la definicion de los datos del generador. Las claves son los ID "
            "tal como estan en ventas.csv (solo aparecen clientes con ventas). 'suma_importe' "
            "suma solo los importes validos (punto decimal, sin pasar por float). "
            "'importes_invalidos' son los que no se pueden convertir (texto, vacio, '12,34,5'). "
            "'ultima_fecha' mira todas las filas del cliente (dd/mm/aaaa). 'fila' es la fila de "
            "datos (1 = primera sin encabezado); 'fila_en_archivo' cuenta el encabezado. "
            "'ids_fuera_de_clientes' son ventas de IDs que no estan en clientes.xlsx."
        ),
        "por_cliente": {
            k: {
                "filas": r["filas"],
                "suma_importe": _cents(r["centavos"]),
                "importes_invalidos": r["invalidos"],
                "articulos_distintos": len(r["arts"]),
                "ultima_fecha": _fecha_ar(r["ultima"]),
            }
            for k, r in sorted(por_cliente.items(), key=lambda kv: _clave(kv[0]))
        },
        "celdas_invalidas": [
            {
                "fila": n,
                "fila_en_archivo": n + 1,
                "id_cliente": v["id"],
                "valor": v["importe_texto"],
            }
            for n, v in enumerate(ventas, start=1)
            if v["centavos"] is None
        ],
        "ids_fuera_de_clientes": [h for h in HUERFANAS_VENTAS],
        "filas_total": len(ventas),
    }

    # --- resumen de cobranzas (por ID real del cliente)
    por_cob: dict[str, dict] = {}
    for cb in cobranzas:
        r = por_cob.setdefault(cb["id"], {"filas": 0, "centavos": 0})
        r["filas"] += 1
        r["centavos"] += cb["centavos"]
    resumen_cobranzas = {
        "_nota": (
            "Calculado desde la definicion de los datos. Las claves son el ID real del cliente "
            "(ya sin ceros ni espacios); en el archivo la clave aparece como '00012' o ' 12 '. "
            "Solo aparecen clientes con cobranzas."
        ),
        "por_cliente": {
            k: {"filas": r["filas"], "suma_monto": _cents(r["centavos"])}
            for k, r in sorted(por_cob.items(), key=lambda kv: _clave(kv[0]))
        },
        "filas_total": len(cobranzas),
    }

    # --- diagnostico
    claves_v = [v["id"] for v in ventas]
    diag_v = _diagnostico(ids, claves_v)
    cuenta_v = {k: claves_v.count(k) for k in set(claves_v)}
    diag_c: dict[str, dict] = {}
    for variante in ("solo_recortar_espacios", "sin_ceros_a_la_izquierda"):
        claves_c = [_normalizar_clave(cb["clave_texto"], variante) for cb in cobranzas]
        dc = _diagnostico(ids, claves_c)
        cuenta_c = {k: claves_c.count(k) for k in set(claves_c)}
        dc["filas_resultado_ventas_detalle_y_cobranzas_detalle"] = sum(
            max(1, cuenta_v.get(k, 0)) * max(1, cuenta_c.get(k, 0)) for k in ids
        )
        dc["filas_resultado_ventas_detalle_y_cobranzas_resumen"] = diag_v["filas_resultado_detalle"]
        dc["filas_resultado_ventas_resumen_y_cobranzas_resumen"] = len(ids)
        diag_c[variante] = dc
    diagnostico = {
        "_nota": (
            "Principal = clientes.xlsx (30 clientes, 1 fila cada uno). Union tipo 'left': cada "
            "cliente aparece al menos una vez. 'filas_resultado_detalle' = suma por cliente de "
            "max(1, filas en el secundario). La limpieza de claves con ceros a la izquierda se "
            "decide en la Fase 4, por eso cobranzas tiene dos variantes: 'solo_recortar_espacios' "
            "(strip: '00012' no coincide con '12') y 'sin_ceros_a_la_izquierda' (strip y sacar "
            "ceros). Las filas 'ventas_detalle_y_cobranzas_detalle' muestran la multiplicacion: "
            "por cliente, filas de ventas por filas de cobranzas. Con las dos en resumen, el "
            "resultado tiene las mismas filas que el principal."
        ),
        "ventas": diag_v,
        "cobranzas": diag_c,
    }

    # --- total por region
    cliente_region = {c["id"]: c["region"] for c in d["clientes"]}
    total_region = dict.fromkeys(REGIONES, 0)
    excluido = 0
    for v in ventas:
        if v["id"] in cliente_region:
            total_region[cliente_region[v["id"]]] += v["centavos"] or 0
        else:
            excluido += v["centavos"] or 0
    por_region = {
        "_nota": (
            "Ventas en resumen (suma de importes validos por cliente) unidas a clientes.xlsx y "
            "sumadas por Region. Los clientes sin ventas suman 0. Las ventas de IDs que no "
            "estan en clientes se excluyen (union left) y se informan aparte."
        ),
        "total_importe_por_region": {r: _cents(c) for r, c in sorted(total_region.items())},
        "clientes_por_region": {
            r: sum(1 for c in d["clientes"] if c["region"] == r) for r in sorted(REGIONES)
        },
        "importe_excluido_ids_fuera_de_clientes": _cents(excluido),
    }
    return {
        "resumen_ventas": resumen_ventas,
        "resumen_cobranzas": resumen_cobranzas,
        "diagnostico": diagnostico,
        "resumen_final_por_region": por_region,
    }


# ---------------------------------------------------------------------------
# Grupo: dificiles
# ---------------------------------------------------------------------------

DUPLICADOS = [
    ("C001", "Ñandú Pérez", "Norte"),
    ("C002", "José Muñoz", "Sur"),
    ("C003", "María Núñez", "Centro"),
    ("C001", "Ñandú Pérez (otra sucursal)", "Sur"),
    ("C004", "Sebastián Peña", "Cuyo"),
    ("C003", "María Núñez (otro domicilio)", "Centro"),
    ("C003", "María Núñez (otro contacto)", "Norte"),
    ("C005", "Inés Ibáñez", "Litoral"),
]
CLAVES_SUCIAS = [
    ("  Peña ", "uno"), ("PEÑA", "dos"), ("peña", "tres"), ("00123", "cuatro"),
    ("123", "cinco"), ("0123", "seis"), ("Ñandú", "siete"), ("ÑANDÚ", "ocho"),
    ("Núñez ", "nueve"), ("núñez", "diez"), ("ABC-01", "once"), ("abc-01 ", "doce"),
]  # fmt: skip
BASURA_ENCABEZADO = ["Código", "Nombre", "Importe"]
BASURA_FILAS = [
    ("00123", "Ñandú Pérez", "1.234,56"),
    ("00124", "José Muñoz", "250,75"),
    ("00125", "María Núñez", "98.765,43"),
    ("00126", "Sebastián Peña", "0,50"),
    ("00127", "Inés Ibáñez", "15.000,25"),
    ("00128", "Ángel Güemes", "310,10"),
]
BASURA_ARRIBA = [["Listado de clientes - Sucursal Córdoba"], ["Emitido: 01/06/2024"], []]


def _normal(k: str) -> str:
    return k.strip().casefold()


def esperados_dificiles() -> dict:
    claves = [k for k, *_ in DUPLICADOS]
    grupos: dict[str, list[str]] = {}
    for k, _ in CLAVES_SUCIAS:
        grupos.setdefault(_normal(k), []).append(k)
    sin_ceros: dict[str, list[str]] = {}
    for k, _ in CLAVES_SUCIAS:
        sin_ceros.setdefault(_normal(k).lstrip("0") or "0", []).append(k)
    basura = {
        "fila_encabezado": len(BASURA_ARRIBA) + 1,
        "filas_basura_antes_del_encabezado": len(BASURA_ARRIBA),
        "columnas": BASURA_ENCABEZADO,
        "filas_de_datos": len(BASURA_FILAS),
    }
    return {
        "_nota": (
            "Calculado desde la definicion de los datos del generador. 'fila_encabezado' es "
            "1-based contando todas las filas del archivo (la fila vacia cuenta). "
            "'grupos_por_clave_normalizada' agrupa las claves con strip + casefold; "
            "'grupos_sin_ceros_a_la_izquierda' ademas saca los ceros a la izquierda."
        ),
        "principal_claves_duplicadas.csv": {
            "fila_encabezado": 1,
            "filas_de_datos": len(DUPLICADOS),
            "columna_clave": "Clave",
            "claves_duplicadas": {
                k: claves.count(k) for k in sorted(set(claves)) if claves.count(k) > 1
            },
            "claves_distintas": len(set(claves)),
        },
        "claves_sucias": {
            "archivos": ["claves_sucias.csv", "claves_sucias.xlsx"],
            "fila_encabezado": 1,
            "filas_de_datos": len(CLAVES_SUCIAS),
            "columna_clave": "Clave",
            "claves_distintas_sin_normalizar": len({k for k, _ in CLAVES_SUCIAS}),
            "grupos_por_clave_normalizada": grupos,
            "grupos_sin_ceros_a_la_izquierda": sin_ceros,
        },
        "con_basura.csv": basura,
        "con_basura.xlsx": basura,
    }


def grupo_dificiles(salida: Path) -> list[str]:
    from openpyxl import Workbook

    def libro() -> Workbook:
        wb = Workbook()
        wb.properties.creator = "make_fixtures"
        wb.properties.created = FECHA_FIJA
        wb.properties.modified = FECHA_FIJA
        return wb

    _escribir_csv(
        ruta_segura(salida, "dificiles/principal_claves_duplicadas.csv"),
        [["Clave", "Nombre", "Región"]] + [list(f) for f in DUPLICADOS],
        ";",
        "utf-8-sig",
    )
    sucias = [["Clave", "Detalle"]] + [list(f) for f in CLAVES_SUCIAS]
    _escribir_csv(ruta_segura(salida, "dificiles/claves_sucias.csv"), sucias, ";", "utf-8-sig")
    wb = libro()
    ws = wb.active
    ws.title = "Claves"
    for fila in sucias:
        ws.append(fila)
    wb.save(ruta_segura(salida, "dificiles/claves_sucias.xlsx"))

    filas_csv = [list(BASURA_ARRIBA[0]) + ["", ""], [BASURA_ARRIBA[1][0], "", ""], ["", "", ""]]
    filas_csv += [BASURA_ENCABEZADO] + [list(f) for f in BASURA_FILAS]
    _escribir_csv(ruta_segura(salida, "dificiles/con_basura.csv"), filas_csv, ";", "utf-8-sig")
    wb = libro()
    ws = wb.active
    ws.title = "Listado"
    ws.cell(row=1, column=1, value=BASURA_ARRIBA[0][0])
    ws.cell(row=2, column=1, value=BASURA_ARRIBA[1][0])
    for c, titulo in enumerate(BASURA_ENCABEZADO, start=1):
        ws.cell(row=4, column=c, value=titulo)
    for r, fila in enumerate(BASURA_FILAS, start=5):
        for c, valor in enumerate(fila, start=1):
            ws.cell(row=r, column=c, value=valor)
    wb.save(ruta_segura(salida, "dificiles/con_basura.xlsx"))

    escribir_json(carpeta_de_esperados(salida) / "dificiles.json", esperados_dificiles())
    return [
        f"dificiles/{n}"
        for n in (
            "principal_claves_duplicadas.csv",
            "claves_sucias.csv",
            "claves_sucias.xlsx",
            "con_basura.csv",
            "con_basura.xlsx",
        )
    ]


# ---------------------------------------------------------------------------
# Grupos grandes (los saltea --sin-grandes): escenario grande y rendimiento
# ---------------------------------------------------------------------------

N_CLIENTES_GRANDE = 200_000
# Ventas por región: la primera es mucho más grande que las demás (total 2.500.000).
VENTAS_POR_REGION = {
    "Norte": 1_400_000,
    "Sur": 275_000,
    "Centro": 275_000,
    "Cuyo": 275_000,
    "Litoral": 275_000,
}
PESOS_CLIENTES = [0.40, 0.15, 0.15, 0.15, 0.15]
FILAS_RENDIMIENTO_CSV = 2_000_000
FILAS_RENDIMIENTO_XLSX = 500_000


def _limites_producto() -> dict:
    """Lee los límites del JSON del producto (es un dato; no se hardcodea acá)."""
    ruta = RAIZ / "src" / "excel_tool" / "limites_default.json"
    return json.loads(ruta.read_text(encoding="utf-8"))


def _importe_ar(centavos: int) -> str:
    return f"{centavos // 100:,}".replace(",", ".") + f",{centavos % 100:02d}"


def estructura_grande() -> dict:
    """Clientes por región y cuáles pueden vender. Todo sale de la semilla, sin archivos."""
    rng = random.Random(SEMILLA + 4)
    regiones = list(VENTAS_POR_REGION)
    region_de = rng.choices(range(len(regiones)), weights=PESOS_CLIENTES, k=N_CLIENTES_GRANDE)
    vendedor_de = [rng.randrange(len(VENDEDORES)) for _ in range(N_CLIENTES_GRANDE)]
    nombre_de = [
        (rng.randrange(len(NOMBRES)), rng.randrange(len(APELLIDOS)))
        for _ in range(N_CLIENTES_GRANDE)
    ]
    # Los clientes de cada región se numeran en orden; el 10% (posición % 10 == 9) nunca vende.
    habilitados: list[list[int]] = [[] for _ in regiones]
    sin_ventas = [0] * len(regiones)
    posicion = [0] * len(regiones)
    for idx, r in enumerate(region_de):
        if posicion[r] % 10 == 9:
            sin_ventas[r] += 1
        else:
            habilitados[r].append(idx)
        posicion[r] += 1
    return {
        "regiones": regiones,
        "region_de": region_de,
        "vendedor_de": vendedor_de,
        "nombre_de": nombre_de,
        "habilitados": habilitados,
        "sin_ventas": sin_ventas,
        "clientes_por_region": posicion,
    }


def _verificar_limites(est: dict) -> dict:
    lim = _limites_producto()["excel"]
    max_filas, comodo = lim["max_rows"], lim["comfortable_rows"]
    regiones = est["regiones"]
    detalle = {r: VENTAS_POR_REGION[r] + est["sin_ventas"][i] for i, r in enumerate(regiones)}
    grande = regiones[0]
    assert detalle[grande] > max_filas, "la region grande tiene que superar el maximo de Excel"
    for r in regiones[1:]:
        assert detalle[r] < comodo, f"la region {r} tiene que quedar bajo el limite comodo"
    return {"max_rows": max_filas, "comfortable_rows": comodo, "detalle": detalle}


def grupo_escenario_grande(salida: Path) -> list[str]:
    est = estructura_grande()
    lim = _verificar_limites(est)
    regiones = est["regiones"]

    with ruta_segura(salida, "grandes/clientes_grande.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as f:
        w = csv.writer(f, delimiter=";", lineterminator="\r\n")
        w.writerow(["ID cliente", "Nombre", "Región", "Vendedor"])
        for idx in range(N_CLIENTES_GRANDE):
            n, a = est["nombre_de"][idx]
            w.writerow(
                [
                    idx + 1,
                    f"{NOMBRES[n]} {APELLIDOS[a]} {idx + 1}",
                    regiones[est["region_de"][idx]],
                    VENDEDORES[est["vendedor_de"][idx]],
                ]
            )

    rng = random.Random(SEMILLA + 5)
    fechas = [_fecha_ar(dt.date(2023, 1, 1) + dt.timedelta(days=d)) for d in range(730)]
    restantes = [VENTAS_POR_REGION[r] for r in regiones]
    emitidas = [0] * len(regiones)
    total = sum(restantes)
    with ruta_segura(salida, "grandes/ventas_grande.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as f:
        w = csv.writer(f, delimiter=";", lineterminator="\r\n")
        w.writerow(["ID cliente", "Fecha", "Artículo", "Cantidad", "Importe"])
        for _ in range(total):
            # Se elige la región en proporción a lo que falta emitir: las filas quedan mezcladas.
            sorteo = rng.randrange(sum(restantes))
            r = 0
            while sorteo >= restantes[r]:
                sorteo -= restantes[r]
                r += 1
            hab = est["habilitados"][r]
            k = emitidas[r]
            idx = hab[k] if k < len(hab) else hab[rng.randrange(len(hab))]
            emitidas[r] += 1
            restantes[r] -= 1
            w.writerow(
                [
                    idx + 1,
                    fechas[rng.randrange(730)],
                    ARTICULOS[rng.randrange(len(ARTICULOS))],
                    rng.randint(1, 20),
                    _importe_ar(rng.randint(1000, 900000)),
                ]
            )

    detalle = lim["detalle"]
    esperado = {
        "_nota": (
            "Calculado desde la definicion de los datos del generador, sin leer los archivos. "
            "'detalle_left' = filas de un resultado en detalle con clientes como principal: "
            "las ventas de la region mas un fila por cada cliente sin ventas. Todos los "
            "clientes habilitados venden al menos una vez; el 10% (posicion % 10 == 9 dentro de "
            "su region) nunca vende. Los limites salen de limites_default.json (Excel)."
        ),
        "limites_excel": {"max_rows": lim["max_rows"], "comfortable_rows": lim["comfortable_rows"]},
        "filas": {
            "clientes_grande.csv": N_CLIENTES_GRANDE,
            "ventas_grande.csv": sum(VENTAS_POR_REGION.values()),
            "rendimiento_2M.csv": FILAS_RENDIMIENTO_CSV,
            "rendimiento_500k.xlsx": FILAS_RENDIMIENTO_XLSX,
        },
        "clientes_por_region": dict(zip(regiones, est["clientes_por_region"], strict=True)),
        "ventas_por_region": dict(VENTAS_POR_REGION),
        "clientes_sin_ventas": {
            "total": sum(est["sin_ventas"]),
            "por_region": dict(zip(regiones, est["sin_ventas"], strict=True)),
        },
        "detalle_left": {
            "total": sum(detalle.values()),
            "por_region": detalle,
        },
        "regiones_que_superan_el_maximo_de_filas_de_excel": [
            r for r in regiones if detalle[r] > lim["max_rows"]
        ],
        "regiones_que_superan_el_limite_comodo_de_excel": [
            r for r in regiones if detalle[r] > lim["comfortable_rows"]
        ],
        "total_detalle_supera_el_maximo_de_excel": sum(detalle.values()) > lim["max_rows"],
    }
    escribir_json(carpeta_de_esperados(salida) / "grandes.json", esperado)
    return ["grandes/clientes_grande.csv", "grandes/ventas_grande.csv"]


def grupo_rendimiento(salida: Path) -> list[str]:
    rng = random.Random(SEMILLA + 6)
    palabras = ["Ñandú", "Cañería", "Lámpara", "Aceite", "Camión", "Árbol", "Tornillo"]
    regiones = ["Norte", "Sur", "Centro", "Cuyo", "Litoral"]
    base = dt.date(2022, 1, 1)

    def fila(i: int) -> list[str]:
        dia = base + dt.timedelta(days=rng.randrange(1000))
        return [
            str(i),
            f"{rng.randrange(100000, 99999999) / 100:.2f}".replace(".", ","),
            _fecha_ar(dia),
            f"{_fecha_ar(dia)} {rng.randrange(24):02d}:{rng.randrange(60):02d}",
            "Sí" if rng.random() < 0.5 else "No",
            f"{palabras[rng.randrange(7)]} {palabras[rng.randrange(7)]}",
            f"{rng.randrange(1, 99999):06d}",
            str(rng.randrange(1000)) if rng.random() < 0.9 else "s/d",
            regiones[rng.randrange(5)],
            "" if rng.random() < 0.3 else f"obs {rng.randrange(1000)}",
        ]

    encabezado = [
        "ID", "Importe", "Fecha", "Fecha y hora", "Activo", "Descripción", "Código", "Mixta",
        "Región", "Observación",
    ]  # fmt: skip
    with ruta_segura(salida, "grandes/rendimiento_2M.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as f:
        w = csv.writer(f, delimiter=";", lineterminator="\r\n")
        w.writerow(encabezado)
        for i in range(1, FILAS_RENDIMIENTO_CSV + 1):
            w.writerow(fila(i))

    from openpyxl import Workbook

    wb = Workbook(write_only=True)
    wb.properties.creator = "make_fixtures"
    wb.properties.created = FECHA_FIJA
    wb.properties.modified = FECHA_FIJA
    ws = wb.create_sheet("Datos")
    ws.append(encabezado)
    for i in range(1, FILAS_RENDIMIENTO_XLSX + 1):
        ws.append([v if v != "" else None for v in fila(i)])
    wb.save(ruta_segura(salida, "grandes/rendimiento_500k.xlsx"))
    return ["grandes/rendimiento_2M.csv", "grandes/rendimiento_500k.xlsx"]


# ---------------------------------------------------------------------------
# Grupo: fase1 (tipos de celda de Excel, CSV raros, archivos rotos) + fase1.json
# ---------------------------------------------------------------------------

FECHAS_F1 = [dt.date(2024, 3, 5), dt.datetime(2023, 12, 31), dt.date(2024, 1, 1)]
FECHAS_HORA_F1 = [
    dt.datetime(2024, 3, 5, 14, 30, 15),
    dt.datetime(2023, 12, 31, 0, 0, 1),
    dt.datetime(2024, 1, 1, 23, 59, 59),
]
ENCABEZADO_TIPOS = [
    "Entero",
    "Decimal",
    "Fecha",
    "FechaHora",
    "Booleano",
    "Formula",
    "Error",
    "Texto",
]


def _fijar_zip(origen: Path, reemplazos: dict[str, list[tuple[str, str]]]) -> None:
    """Reescribe un .xlsx con fechas fijas en el zip y reemplazos de texto en sus partes
    (openpyxl no guarda el valor calculado de las fórmulas: se lo agregamos a mano)."""
    with zipfile.ZipFile(origen) as z:
        partes = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(origen, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, datos in partes:
            for viejo, nuevo in reemplazos.get(nombre, []):
                texto = datos.decode("utf-8")
                if viejo not in texto:
                    raise ValueError(f"No está {viejo!r} en {nombre}")
                datos = texto.replace(viejo, nuevo).encode("utf-8")
            info = zipfile.ZipInfo(nombre, date_time=(2024, 6, 1, 12, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, datos)


def generar_excel_tipos_xlsx(destino: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill

    wb = Workbook()
    wb.properties.creator = "make_fixtures"
    wb.properties.created = FECHA_FIJA
    wb.properties.modified = FECHA_FIJA
    ws = wb.active
    ws.title = "Tipos"
    ws.append(ENCABEZADO_TIPOS)
    ws.append([100, 1234.56, FECHAS_F1[0], FECHAS_HORA_F1[0], True, "=A2+1", "#N/A", "a"])
    ws.append([0, 0.5, FECHAS_F1[1], FECHAS_HORA_F1[1], False, "=A3+1", None, "b"])
    ws.append([-5, -2.25, FECHAS_F1[2], FECHAS_HORA_F1[2], True, "=NA()", None, "c"])
    ws.append([1048577, 0.001, None, None, None, None, None, None])
    for fila in (2, 3, 4):
        ws.cell(fila, 3).number_format = "dd/mm/yyyy"
        ws.cell(fila, 4).number_format = "dd/mm/yyyy hh:mm:ss"
    ws.merge_cells("H2:H3")  # Texto: "a" ocupa 2 filas (la segunda queda vacía)
    ws.merge_cells("G3:G4")  # Error: dos celdas vacías combinadas
    ws2 = wb.create_sheet("Con vacias")
    ws2.append(["A", "B"])
    ws2.append(["x", 1])
    ws2.append(["y", 2])
    for fila in (4, 5, 6):  # filas "con formato" pero sin datos, al final
        for col in (1, 2):
            ws2.cell(fila, col).fill = PatternFill("solid", fgColor="DDDDDD")
    wb.save(destino)
    _fijar_zip(
        destino,
        {
            "xl/worksheets/sheet1.xml": [
                ('<c r="F2"><f>A2+1</f><v /></c>', '<c r="F2"><f>A2+1</f><v>101</v></c>'),
                ('<c r="F3"><f>A3+1</f><v /></c>', '<c r="F3"><f>A3+1</f><v>1</v></c>'),
                (
                    '<c r="F4"><f>NA()</f><v /></c>',
                    '<c r="F4" t="e"><f>NA()</f><v>#N/A</v></c>',
                ),
            ]
        },
    )


def generar_excel_tipos_xls(destino: Path) -> None:
    import xlwt

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Tipos")
    estilo_fecha = xlwt.easyxf(num_format_str="DD/MM/YYYY")
    estilo_hora = xlwt.easyxf(num_format_str="DD/MM/YYYY HH:MM:SS")
    for c, titulo in enumerate(["Entero", "Decimal", "Fecha", "FechaHora", "Booleano"]):
        ws.write(0, c, titulo)
    filas = [
        (100, 1234.56, dt.date(2024, 3, 5), dt.datetime(2024, 3, 5, 14, 30, 15), True),
        (0, 0.5, dt.date(2023, 12, 31), dt.datetime(2023, 12, 31, 0, 0, 1), False),
        (-5, -2.25, dt.date(2024, 1, 1), dt.datetime(2024, 1, 1, 23, 59, 59), True),
    ]
    for r, (e, d, f, fh, b) in enumerate(filas, start=1):
        ws.write(r, 0, e)
        ws.write(r, 1, d)
        ws.write(r, 2, f, estilo_fecha)
        ws.write(r, 3, fh, estilo_hora)
        ws.write(r, 4, b)
    wb.save(str(destino))


def esperados_fase1() -> dict:
    """Todo escrito a mano, a partir de las reglas del contrato (no se lee ningún archivo)."""
    cols_tipos = {
        "Entero": "entero",
        "Decimal": "decimal",
        "Fecha": "fecha",
        "FechaHora": "fecha_hora",
        "Booleano": "si_no",
        "Formula": "entero",
        "Error": "vacia",
        "Texto": "texto",
    }
    return {
        "_nota": (
            "Valores como los devuelve read_table (todo texto): enteros sin '.0' ('100'), "
            "decimales con punto ('1234.56'), fechas 'dd/mm/aaaa', fecha-hora "
            "'dd/mm/aaaa hh:mm:ss', booleanos 'VERDADERO'/'FALSO', vacio ''. Errores de "
            "Excel (#N/A): calamine 0.9.1 los devuelve como '' (no como texto): el esperado "
            "es ''; el contrato dice '#N/A'. "
            "'columnas' y 'filas' son lo que debe devolver "
            "read_table con header_row=None; 'tipos' es el ColumnType que detect_types debe "
            "dar por columna. 'error' es la clase de core/errors que debe levantar la lectura. "
            "Formulas: openpyxl no guarda el valor calculado, asi que el generador lo agrega "
            "al XML (101, 1 y #N/A); el esperado es ese valor. Celdas combinadas: el valor "
            "queda en la primera celda y las demas leen vacio (''). 'filas_vacias_final' de "
            "la hoja 'Con vacias': filas en blanco con formato al final, que read_table "
            "descarta (calamine las puede devolver o no; el archivo las tiene)."
        ),
        "excel_tipos.xlsx": {
            "hoja_principal": "Tipos",
            "columnas": ENCABEZADO_TIPOS,
            "filas": [
                [
                    "100",
                    "1234.56",
                    "05/03/2024",
                    "05/03/2024 14:30:15",
                    "VERDADERO",
                    "101",
                    "",
                    "a",
                ],
                ["0", "0.5", "31/12/2023", "31/12/2023 00:00:01", "FALSO", "1", "", ""],
                ["-5", "-2.25", "01/01/2024", "01/01/2024 23:59:59", "VERDADERO", "", "", "c"],
                ["1048577", "0.001", "", "", "", "", "", ""],
            ],
            "tipos": cols_tipos,
            "celdas_combinadas": ["H2:H3", "G3:G4"],
            "hojas": {
                "Tipos": {"filas": 4, "columnas": 8},
                "Con vacias": {"filas": 2, "columnas": 2, "filas_vacias_final": 3},
            },
            "hoja_con_vacias": {
                "columnas": ["A", "B"],
                "filas": [["x", "1"], ["y", "2"]],
            },
        },
        "excel_tipos.xls": {
            "hoja_principal": "Tipos",
            "columnas": ["Entero", "Decimal", "Fecha", "FechaHora", "Booleano"],
            "filas": [
                ["100", "1234.56", "05/03/2024", "05/03/2024 14:30:15", "VERDADERO"],
                ["0", "0.5", "31/12/2023", "31/12/2023 00:00:01", "FALSO"],
                ["-5", "-2.25", "01/01/2024", "01/01/2024 23:59:59", "VERDADERO"],
            ],
            "tipos": {
                "Entero": "entero",
                "Decimal": "decimal",
                "Fecha": "fecha",
                "FechaHora": "fecha_hora",
                "Booleano": "si_no",
            },
        },
        "csv/pipe.csv": {
            "separador": "|",
            "encoding": "utf-8",
            "columnas": ["Código", "Nombre", "Importe"],
            "filas": [
                ["00123", "Ñandú Pérez", "1234.56"],
                ["00124", "José Muñoz", "250.75"],
                ["00125", "María Núñez", "-5"],
            ],
            "tipos": {"Código": "texto", "Nombre": "texto", "Importe": "decimal"},
        },
        "csv/estilo_punto.csv": {
            "separador": ",",
            "encoding": "utf-8",
            "columnas": ["Producto", "Precio"],
            "filas": [
                ["Tuerca", "1,234.56"],
                ["Tornillo", "250.75"],
                ["Arandela", "0.5"],
            ],
            "tipos": {"Producto": "texto", "Precio": "decimal"},
            "decimal_style": ".",
        },
        "csv/solo_encabezado.csv": {
            "separador": ";",
            "encoding": "utf-8",
            "columnas": ["Código", "Nombre", "Importe"],
            "filas": [],
            "tipos": {"Código": "vacia", "Nombre": "vacia", "Importe": "vacia"},
        },
        "csv/vacio.csv": {"error": "ArchivoIlegible", "bytes": 0},
        "danado.xlsx": {"error": "ArchivoIlegible"},
    }


def grupo_fase1(salida: Path) -> list[str]:
    generar_excel_tipos_xlsx(ruta_segura(salida, "fase1/excel_tipos.xlsx"))
    generar_excel_tipos_xls(ruta_segura(salida, "fase1/excel_tipos.xls"))
    _escribir_csv(
        ruta_segura(salida, "fase1/csv/pipe.csv"),
        [
            ["Código", "Nombre", "Importe"],
            ["00123", "Ñandú Pérez", "1234.56"],
            ["00124", "José Muñoz", "250.75"],
            ["00125", "María Núñez", "-5"],
        ],
        "|",
        "utf-8",
    )
    _escribir_csv(
        ruta_segura(salida, "fase1/csv/estilo_punto.csv"),
        [
            ["Producto", "Precio"],
            ["Tuerca", "1,234.56"],
            ["Tornillo", "250.75"],
            ["Arandela", "0.5"],
        ],
        ",",
        "utf-8",
    )
    _escribir_csv(
        ruta_segura(salida, "fase1/csv/solo_encabezado.csv"),
        [["Código", "Nombre", "Importe"]],
        ";",
        "utf-8",
    )
    ruta_segura(salida, "fase1/csv/vacio.csv").write_bytes(b"")
    ruta_segura(salida, "fase1/danado.xlsx").write_bytes(
        "Esto no es un Excel, es texto suelto: ñandú.\n".encode() * 20
    )
    escribir_json(carpeta_de_esperados(salida) / "fase1.json", esperados_fase1())
    return [
        "fase1/excel_tipos.xlsx",
        "fase1/excel_tipos.xls",
        "fase1/csv/pipe.csv",
        "fase1/csv/estilo_punto.csv",
        "fase1/csv/solo_encabezado.csv",
        "fase1/csv/vacio.csv",
        "fase1/danado.xlsx",
    ]


# ---------------------------------------------------------------------------
# Registro de grupos (F0-21 y F0-22 suman acá)
# ---------------------------------------------------------------------------

# (nombre, función, es_grande)
GRUPOS: list[tuple[str, Callable[[Path], list[str]], bool]] = [
    ("formatos", grupo_formatos, False),
    ("csv", grupo_csv, False),
    ("explorar", grupo_explorar, False),
    ("escenario_3", grupo_escenario_3, False),
    ("dificiles", grupo_dificiles, False),
    ("fase1", grupo_fase1, False),
    ("escenario_grande", grupo_escenario_grande, True),
    ("rendimiento", grupo_rendimiento, True),
]


def _pico_memoria_mb() -> float | None:
    try:
        import resource  # solo existe en Linux/macOS
    except ImportError:
        return None
    pico = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return pico / (1024 if sys.platform != "darwin" else 1024 * 1024)


def generar(salida: Path, sin_grandes: bool = False) -> list[str]:
    salida.mkdir(parents=True, exist_ok=True)
    creados: list[str] = []
    marca = salida / "grandes" / "LISTO"
    hubo_grandes = False
    for nombre, funcion, grande in GRUPOS:
        if grande and sin_grandes:
            print(f"Salteando el grupo '{nombre}' (archivos grandes).")
            continue
        if grande and not hubo_grandes:
            marca.unlink(missing_ok=True)  # hasta terminar, los grandes no estan "listos"
            hubo_grandes = True
        print(f"Generando el grupo '{nombre}'...")
        inicio = time.perf_counter()
        archivos = funcion(salida)
        if grande:
            for rel in archivos:
                mb = (salida / rel).stat().st_size / (1024 * 1024)
                print(f"  {rel}: {mb:,.1f} MB")
            print(f"  tardo {time.perf_counter() - inicio:,.1f} s")
        creados += archivos
    (salida / ARCHIVO_VERSION).write_text(huella_generador() + "\n", encoding="utf-8", newline="\n")
    if hubo_grandes:
        marca.write_text(huella_generador() + "\n", encoding="utf-8", newline="\n")
    pico = _pico_memoria_mb()
    if hubo_grandes and pico is not None:
        print(f"Pico de memoria aproximado: {pico:,.0f} MB")
    return creados


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Genera los archivos de prueba.")
    ap.add_argument("--sin-grandes", action="store_true", help="saltea los archivos grandes")
    ap.add_argument("--salida", type=Path, default=SALIDA_DEFECTO, help="carpeta de salida")
    args = ap.parse_args(argv)
    creados = generar(args.salida, args.sin_grandes)
    print(f"Listo: {len(creados)} archivos en {args.salida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

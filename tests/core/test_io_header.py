from excel_tool.core.io.header import column_letter, detect_header_row, normalize_column_names


def test_encabezado_en_primera_fila():
    filas = [["A", "B", "C"], ["1", "2", "3"], ["4", "5", "6"]]
    assert detect_header_row(filas) == 0


def test_encabezado_despues_de_basura():
    filas = [
        ["Listado de clientes", "", ""],
        ["Emitido: 01/06/2024", "", ""],
        ["", "", ""],
        ["Código", "Nombre", "Importe"],
        ["00123", "Ñandú", "1,5"],
        ["00124", "José", "2,5"],
    ]
    assert detect_header_row(filas) == 3


def test_encabezado_con_una_columna_sin_nombre():
    filas = [["A", "", "C", "D"], ["1", "2", "3", "4"], ["5", "6", "7", "8"]]
    assert detect_header_row(filas) == 0


def test_sin_encabezado_claro_usa_fila_cero():
    assert detect_header_row([]) == 0
    assert detect_header_row([[""], [""]]) == 0
    assert detect_header_row([["solo"], ["una"], ["columna"]]) == 0


def test_solo_mira_30_filas():
    filas = [["x", "", ""]] * 40 + [["A", "B", "C"], ["1", "2", "3"]]
    assert detect_header_row(filas) == 0


def test_letras_de_columna():
    assert [column_letter(i) for i in (0, 5, 25, 26, 27, 701, 702)] == [
        "A",
        "F",
        "Z",
        "AA",
        "AB",
        "ZZ",
        "AAA",
    ]


def test_nombres_normales_y_recorte():
    finales, renamed, avisos = normalize_column_names(["  A ", "B"])
    assert finales == ["A", "B"]
    assert renamed == {}
    assert avisos == []


def test_nombres_vacios():
    finales, renamed, avisos = normalize_column_names(["A", "", "C", "", "  ", "F"])
    assert finales == ["A", "Columna B", "C", "Columna D", "Columna E", "F"]
    assert set(renamed) == {"Columna B", "Columna D", "Columna E"}
    assert len(avisos) == 1
    assert "3 columnas sin nombre" in avisos[0]
    assert "Columna B, Columna D y Columna E" in avisos[0]


def test_nombre_vacio_singular():
    _, _, avisos = normalize_column_names(["A", ""])
    assert avisos == ["Hay 1 columna sin nombre: se llamó Columna B."]


def test_nombres_repetidos():
    finales, renamed, avisos = normalize_column_names(["Id", "Monto", "Id", "Id", "Monto"])
    assert finales == ["Id", "Monto", "Id (2)", "Id (3)", "Monto (2)"]
    assert renamed == {"Id (2)": "Id", "Id (3)": "Id", "Monto (2)": "Monto"}
    assert len(avisos) == 1
    assert "5 columnas con el mismo nombre" in avisos[0]


def test_repetido_no_choca_con_existente():
    finales, _, _ = normalize_column_names(["Id", "Id", "Id (2)"])
    assert len(set(finales)) == 3

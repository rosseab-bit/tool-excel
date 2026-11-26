"""Grilla de datos paginada y reutilizable.

Recibe un proveedor de páginas: una función ``(página) -> (DataFrame, total de filas)``
con la página en base 0. La grilla nunca guarda más de una página en el widget y no
ordena ni filtra: avisa con ``on_sort`` y quien la usa entrega un proveedor nuevo.
"""

import logging
import tkinter as tk
import tkinter.font as tkfont
from collections.abc import Callable
from tkinter import messagebox, ttk
from typing import Any, Protocol

log = logging.getLogger(__name__)

COLUMNA_FILA = "N.º fila"
TAMANO_PAGINA = 1000
MAX_CARACTERES = 120
ANCHO_MIN = 60
ANCHO_MAX = 320
FILAS_PARA_ANCHO = 100

TEXTO_SIN_DATOS = "Sin filas para mostrar"
TEXTO_ERROR = "No se pudieron mostrar los datos. Probá de nuevo."


class ProveedorDePaginas(Protocol):
    def __call__(self, pagina: int) -> tuple[Any, int]:
        """Devuelve (DataFrame de la página, total de filas). Página en base 0."""
        ...


def formatear_numero(n: int) -> str:
    """1234567 -> '1.234.567'."""
    return f"{n:,}".replace(",", ".")


def texto_celda(valor: Any) -> str:
    """Vacío en blanco; saltos de línea a espacio; textos largos recortados."""
    if valor is None or valor != valor:
        return ""
    texto = str(valor).replace("\r", " ").replace("\n", " ")
    if len(texto) > MAX_CARACTERES:
        texto = texto[: MAX_CARACTERES - 1] + "…"
    return texto


class Grilla(ttk.Frame):
    def __init__(
        self,
        master: tk.Misc,
        proveedor: ProveedorDePaginas | None = None,
        page_size: int = TAMANO_PAGINA,
        on_sort: Callable[[str, bool], None] | None = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(master, **kwargs)
        self.page_size = max(1, int(page_size))
        self.on_sort = on_sort
        self._proveedor = proveedor
        self._pagina = 0
        self._total = 0
        self._nombres: list[str] = []
        self._ids: dict[str, str] = {}
        self._ocultas: set[str] = set()
        self._orden: tuple[str, bool] | None = None
        self._fuente = tkfont.nametofont("TkDefaultFont")

        barra = ttk.Frame(self)
        barra.pack(side="bottom", fill="x", pady=(4, 0))
        self._btn_primera = ttk.Button(barra, text="« Primera", width=9, command=self.primera)
        self._btn_anterior = ttk.Button(barra, text="‹ Anterior", width=10, command=self.anterior)
        self._btn_siguiente = ttk.Button(
            barra, text="Siguiente ›", width=10, command=self.siguiente
        )
        self._btn_ultima = ttk.Button(barra, text="Última »", width=9, command=self.ultima)
        self._posicion = tk.StringVar(master=self, value="")
        for b in (self._btn_primera, self._btn_anterior):
            b.pack(side="left", padx=2)
        ttk.Label(barra, textvariable=self._posicion).pack(side="left", padx=10)
        for b in (self._btn_ultima, self._btn_siguiente):
            b.pack(side="right", padx=2)

        cuerpo = ttk.Frame(self)
        cuerpo.pack(side="top", fill="both", expand=True)
        cuerpo.rowconfigure(0, weight=1)
        cuerpo.columnconfigure(0, weight=1)
        self.arbol = ttk.Treeview(cuerpo, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(cuerpo, orient="vertical", command=self.arbol.yview)
        hsb = ttk.Scrollbar(cuerpo, orient="horizontal", command=self.arbol.xview)
        self.arbol.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.arbol.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        for evento in ("<Button-3>", "<Button-2>"):
            self.arbol.bind(evento, self._menu_columnas, add="+")

        self._actualizar_controles()

    # ---- API pública ----
    @property
    def pagina(self) -> int:
        return self._pagina

    @property
    def total_filas(self) -> int:
        return self._total

    @property
    def cantidad_paginas(self) -> int:
        return max(1, -(-self._total // self.page_size))

    @property
    def texto_posicion(self) -> str:
        return self._posicion.get()

    @property
    def columnas(self) -> list[str]:
        return list(self._nombres)

    @property
    def columnas_visibles(self) -> list[str]:
        return [n for n in self._nombres if n not in self._ocultas]

    def set_proveedor(self, proveedor: ProveedorDePaginas | None, pagina: int = 0) -> None:
        """Cambia el proveedor (por ejemplo tras ordenar o filtrar) y va a la página dada."""
        self._proveedor = proveedor
        self.ir_a(pagina)

    def refrescar(self) -> None:
        self.ir_a(self._pagina)

    def ir_a(self, pagina: int) -> None:
        if self._proveedor is None:
            self._vaciar()
            return
        try:
            df, total = self._proveedor(max(0, pagina))
        except Exception:
            log.exception("Falló el proveedor de páginas de la grilla")
            messagebox.showerror("Datos", TEXTO_ERROR, parent=self)
            return
        self._total = int(total)
        ultima = self.cantidad_paginas - 1
        if pagina > ultima:
            try:
                df, self._total = self._proveedor(ultima)
            except Exception:
                log.exception("Falló el proveedor de páginas de la grilla")
                messagebox.showerror("Datos", TEXTO_ERROR, parent=self)
                return
        self._pagina = max(0, min(pagina, ultima))
        self._mostrar(df)
        self._actualizar_controles()

    def primera(self) -> None:
        self.ir_a(0)

    def anterior(self) -> None:
        if self._pagina > 0:
            self.ir_a(self._pagina - 1)

    def siguiente(self) -> None:
        if self._pagina < self.cantidad_paginas - 1:
            self.ir_a(self._pagina + 1)

    def ultima(self) -> None:
        self.ir_a(self.cantidad_paginas - 1)

    def ocultar_columna(self, nombre: str) -> None:
        if nombre == COLUMNA_FILA or nombre not in self._ids:
            return
        self._ocultas.add(nombre)
        self._aplicar_visibles()

    def mostrar_columna(self, nombre: str) -> None:
        self._ocultas.discard(nombre)
        self._aplicar_visibles()

    def mostrar_todas(self) -> None:
        self._ocultas.clear()
        self._aplicar_visibles()

    def filas_mostradas(self) -> int:
        return len(self.arbol.get_children())

    # ---- internos ----
    def _vaciar(self) -> None:
        self.arbol.delete(*self.arbol.get_children())
        self._total = 0
        self._pagina = 0
        self._actualizar_controles()

    def _mostrar(self, df: Any) -> None:
        nombres = [str(c) for c in df.columns]
        if COLUMNA_FILA in nombres:  # siempre a la izquierda
            nombres.remove(COLUMNA_FILA)
            nombres.insert(0, COLUMNA_FILA)
            orden_df = [COLUMNA_FILA] + [c for c in df.columns if str(c) != COLUMNA_FILA]
        else:
            orden_df = list(df.columns)
        self.arbol.delete(*self.arbol.get_children())
        if nombres != self._nombres:
            self._nombres = nombres
            self._ocultas &= set(nombres)
            if self._orden and self._orden[0] not in nombres:
                self._orden = None
            self._ids = {n: f"c{i}" for i, n in enumerate(nombres)}
            self.arbol.configure(displaycolumns="#all")  # evita ids viejos al cambiar columnas
            self.arbol.configure(columns=[self._ids[n] for n in nombres])
        filas = [[texto_celda(v) for v in fila] for fila in df[orden_df].itertuples(index=False)]
        for n in nombres:
            self._configurar_columna(n, filas, nombres.index(n))
        for fila in filas:
            self.arbol.insert("", "end", values=fila)
        self._aplicar_visibles()

    def _configurar_columna(self, nombre: str, filas: list[list[str]], pos: int) -> None:
        cid = self._ids[nombre]
        titulo = nombre
        if self._orden and self._orden[0] == nombre:
            titulo += " ▲" if self._orden[1] else " ▼"
        medir = [titulo] + [f[pos] for f in filas[:FILAS_PARA_ANCHO]]
        ancho = max(self._fuente.measure(t) for t in medir) + 24
        ancho = max(ANCHO_MIN, min(ANCHO_MAX, ancho))
        anclaje = "e" if nombre == COLUMNA_FILA else "w"
        self.arbol.heading(cid, text=titulo, command=lambda n=nombre: self._clic_encabezado(n))
        self.arbol.column(cid, width=ancho, minwidth=ANCHO_MIN, stretch=False, anchor=anclaje)

    def _aplicar_visibles(self) -> None:
        visibles = [self._ids[n] for n in self._nombres if n not in self._ocultas]
        self.arbol.configure(displaycolumns=visibles if visibles else "#all")

    def _clic_encabezado(self, nombre: str) -> None:
        if nombre == COLUMNA_FILA:
            return
        ascendente = not (self._orden and self._orden[0] == nombre and self._orden[1])
        self._orden = (nombre, ascendente)
        for n in self._nombres:
            if n != COLUMNA_FILA:
                titulo = n + ((" ▲" if ascendente else " ▼") if n == nombre else "")
                self.arbol.heading(self._ids[n], text=titulo)
        if self.on_sort is not None:
            self.on_sort(nombre, ascendente)

    def _menu_columnas(self, evento: tk.Event) -> None:
        if self.arbol.identify_region(evento.x, evento.y) != "heading" or not self._nombres:
            return
        menu = tk.Menu(self, tearoff=False)
        for n in self._nombres:
            if n == COLUMNA_FILA:
                continue
            var = tk.BooleanVar(master=menu, value=n not in self._ocultas)
            menu.add_checkbutton(
                label=n if len(n) <= 40 else n[:39] + "…",
                variable=var,
                command=lambda n=n, v=var: (
                    self.mostrar_columna(n) if v.get() else self.ocultar_columna(n)
                ),
            )
        menu.add_separator()
        menu.add_command(label="Mostrar todas las columnas", command=self.mostrar_todas)
        try:
            menu.tk_popup(evento.x_root, evento.y_root)
        finally:
            menu.grab_release()

    def _actualizar_controles(self) -> None:
        if self._total == 0:
            self._posicion.set(TEXTO_SIN_DATOS)
        else:
            desde = self._pagina * self.page_size + 1
            hasta = min(self._total, (self._pagina + 1) * self.page_size)
            d, h, t = (formatear_numero(x) for x in (desde, hasta, self._total))
            self._posicion.set(f"Filas {d} a {h} de {t}")
        hay_ant = self._pagina > 0
        hay_sig = self._pagina < self.cantidad_paginas - 1
        for b, ok in (
            (self._btn_primera, hay_ant),
            (self._btn_anterior, hay_ant),
            (self._btn_siguiente, hay_sig),
            (self._btn_ultima, hay_sig),
        ):
            b.state(["!disabled"] if ok else ["disabled"])

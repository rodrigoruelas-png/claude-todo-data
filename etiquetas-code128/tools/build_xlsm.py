"""Construye Etiqueta_Orden_Operador_Code128_v2_vertical.xlsm.

    python3 tools/build_xlsm.py [salida.xlsm]

Hojas:
  ETIQUETA  captura de datos + etiqueta 4 x 6 in girada como Label Matrix (lo que se imprime)
  LISTA     varias ordenes para imprimir en ciclo
  CONFIG    fuente de codigo de barras, prueba e instrucciones
"""

import os
import sys

import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import code128  # noqa: E402
import vba_project  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "Etiqueta_Orden_Operador_Code128_v2_vertical.xlsm")

BARCODE_FONT = "Libre Barcode 128"
BARCODE_SIZE = 35.5          # 0.03 em por modulo -> 3 puntos por modulo a 203 dpi
MAPEO_DEFAULT = "LIBRE / IDAUTOMATION"
MAPEOS = [MAPEO_DEFAULT, "GRANDZEBU (Code 128)"]
TEXT_FONT = "Arial"

# La etiqueta se diseña "acostada" (6 x 4 in): 47 columnas de 0.125 in y las filas
# de LABEL_ROWS. Luego se gira 90 grados (texto de abajo hacia arriba) y se pone en
# la hoja de pie (4 x 6 in), igual que la plantilla ORDEN OPERADOR de Label Matrix.
# Asi se imprime vertical y el driver de la Zebra no tiene que girar nada.
LABEL_COLS = 47
LROW0 = 1                    # la etiqueta girada empieza en la fila 2
LCOL0 = 1                    # y en la columna B
ROW_PT = 9                   # cada columna de diseño = una fila de 9 pt (0.125 in)
LABEL_ROWS = [               # alto de cada fila de diseño (= ancho de columna), en pt
    10, 23, 11, 4,           # 0-3   JOB / OB / QTY
    10, 23, 11, 3,           # 4-7   P/N
    10, 13, 3,               # 8-10  HEAT CODE / HEAT LOT
    17, 10, 10, 10, 10, 3,   # 11-16 tabla ASSY y FECHA DE CALIDAD
    12, 9, 9, 9, 9, 9, 9, 9, 9,  # 17-25 TEST DATA y QA
    9,                       # 26    pie
]
assert sum(LABEL_ROWS) <= 279, sum(LABEL_ROWS)     # 3.875 in

THIN, MEDIUM = 1, 2


class Canvas:
    """Guarda estilos por celda y celdas combinadas; luego las escribe."""

    def __init__(self, wb, ws):
        self.wb, self.ws = wb, ws
        self.props = {}
        self.merges = []
        self.cells = {}
        self.fmt_cache = {}

    def _p(self, r, c):
        return self.props.setdefault((r, c), {})

    def style(self, r1, c1, r2, c2, **kw):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self._p(r, c).update(kw)

    def border(self, r, c, side, weight):
        p = self._p(r, c)
        p[side] = max(p.get(side, 0), weight)

    def box(self, r1, c1, r2, c2, weight=MEDIUM):
        for c in range(c1, c2 + 1):
            self.border(r1, c, "top", weight)
            self.border(r2, c, "bottom", weight)
        for r in range(r1, r2 + 1):
            self.border(r, c1, "left", weight)
            self.border(r, c2, "right", weight)

    def hline(self, r, c1, c2, weight=THIN, side="bottom"):
        for c in range(c1, c2 + 1):
            self.border(r, c, side, weight)

    def vline(self, c, r1, r2, weight=THIN, side="left"):
        for r in range(r1, r2 + 1):
            self.border(r, c, side, weight)

    def put(self, r1, c1, r2, c2, value, formula=False, cached=None, **kw):
        self.style(r1, c1, r2, c2, **kw)
        if (r1, c1) != (r2, c2):
            self.merges.append((r1, c1, r2, c2))
        self.cells[(r1, c1)] = (value, formula, cached)

    def fmt(self, r, c):
        p = self.props.get((r, c), {})
        key = tuple(sorted(p.items()))
        if key not in self.fmt_cache:
            self.fmt_cache[key] = self.wb.add_format(dict(p))
        return self.fmt_cache[key]

    def flush(self):
        merged = set()
        for r1, c1, r2, c2 in self.merges:
            value, formula, cached = self.cells[(r1, c1)]
            self.ws.merge_range(r1, c1, r2, c2, "", self.fmt(r1, c1))
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    merged.add((r, c))
                    if (r, c) != (r1, c1):
                        self.ws.write_blank(r, c, None, self.fmt(r, c))
        for (r, c), (value, formula, cached) in self.cells.items():
            if formula:
                self.ws.write_formula(r, c, value, self.fmt(r, c), cached if cached is not None else "")
            elif value is None or value == "":
                self.ws.write_blank(r, c, None, self.fmt(r, c))
            else:
                self.ws.write(r, c, value, self.fmt(r, c))
        for (r, c) in self.props:
            if (r, c) not in merged and (r, c) not in self.cells:
                self.ws.write_blank(r, c, None, self.fmt(r, c))


# Giro de 90 grados contra reloj: lo de arriba queda a la izquierda, lo de la
# izquierda queda abajo. (fila k, columna c) del diseño -> (fila, columna) de la hoja.
TURN_SIDE = {"top": "left", "bottom": "right", "left": "bottom", "right": "top"}
TURN_ALIGN = {"left": "bottom", "center": "vcenter", "right": "top"}
TURN_VALIGN = {"top": "left", "vcenter": "center", "bottom": "right"}


def turn_cell(k, c):
    return LROW0 + LABEL_COLS - c, LCOL0 + k


def turn_props(p):
    q = {}
    for key, val in p.items():
        if key in TURN_SIDE:
            q[TURN_SIDE[key]] = val
        elif key not in ("align", "valign"):
            q[key] = val
    q["valign"] = TURN_ALIGN[p.get("align", "left")]
    q["align"] = TURN_VALIGN[p.get("valign", "bottom")]
    q["rotation"] = 90
    return q


def turn_into(src, dst):
    """Copia el diseño acostado (src) girado a la hoja (dst). Regresa los nombres."""
    top_left = {}
    for (k, c), p in src.props.items():
        dst.props[turn_cell(k, c)] = turn_props(p)
    for k1, c1, k2, c2 in src.merges:
        r1, s1 = turn_cell(k1, c2)
        r2, s2 = turn_cell(k2, c1)
        dst.merges.append((r1, s1, r2, s2))
        top_left[(k1, c1)] = (r1, s1)
    for (k, c), cell in src.cells.items():
        dst.cells[top_left.get((k, c), turn_cell(k, c))] = cell
    return {name: top_left.get(kc, turn_cell(*kc)) for name, kc in src.names.items()}


def label_layout(cv):
    """Diseño de ESTAMPA UNICA / ORDEN OPERADOR, acostado (6 x 4)."""

    def R(k):
        return k

    def C(c):
        return c

    def put(k1, c1, k2, c2, value, **kw):
        kw.setdefault("font_name", TEXT_FONT)
        cv.put(R(k1), C(c1), R(k2), C(c2), value, **kw)

    cap = dict(bold=True, font_size=7, valign="bottom")
    hr = dict(bold=True, font_size=8, align="center", valign="top")
    bc = dict(font_name=BARCODE_FONT, font_size=BARCODE_SIZE, align="center",
              valign="top", shrink=True)
    small = dict(font_size=6, valign="vcenter")
    black = dict(bg_color="#000000", font_color="#FFFFFF", bold=True, valign="vcenter")

    # --- W.O. / JOB, OB y QTY TOTAL ------------------------------------
    for (c1, c2, title, field, bc_name) in [
        (1, 15, "W.O. / JOB", "In_JOB", "BC_JOB"),
        (17, 31, "OB", "In_OB", "BC_OB"),
        (33, 47, "QTY TOTAL", "In_QTY", "BC_QTY"),
    ]:
        put(0, c1, 0, c2, title, **cap)
        put(1, c1 + 1, 1, c2 - 1, "=IFERROR(CODE128(UPPER(%s),cfgMapeo),\"\")" % field,
            formula=True, **bc)
        cv.names[bc_name] = (R(1), C(c1 + 1))
        put(2, c1, 2, c2, '=UPPER(TRIM(%s))' % field, formula=True, **hr)
    cv.hline(R(2), C(1), C(15))
    cv.hline(R(2), C(17), C(31))
    cv.box(R(0), C(33), R(2), C(47))

    # --- P/N (MODELO) ----------------------------------------------------
    put(4, 1, 4, 47, "P/N (MODELO)", **cap)
    put(5, 3, 5, 45, "=IFERROR(CODE128(UPPER(In_PN),cfgMapeo),\"\")", formula=True, **bc)
    cv.names["BC_PN"] = (R(5), C(3))
    put(6, 1, 6, 47, "=UPPER(TRIM(In_PN))", formula=True, **hr)
    cv.hline(R(6), C(1), C(47))

    # --- HEAT CODE / HEAT LOT -------------------------------------------
    put(8, 1, 8, 24, "HEAT CODE", **cap)
    put(8, 26, 8, 47, "HEAT LOT", **cap)
    put(9, 1, 9, 24, "=UPPER(TRIM(In_HEATCODE))", formula=True, bold=True, font_size=8,
        valign="bottom")
    put(9, 26, 9, 47, "=UPPER(TRIM(In_HEATLOT))", formula=True, bold=True, font_size=8,
        valign="bottom")
    cv.hline(R(9), C(1), C(24))
    cv.hline(R(9), C(26), C(47))

    # --- Tabla ASSY / TEST / NET ----------------------------------------
    heads = ["ASSY\nDATE", "ASSY\nBADGE", "ASSY\nQTY", "TEST\nDATE", "TEST\nBADGE", "NET\nQTY"]
    for j, h in enumerate(heads):
        c1, c2 = 1 + 4 * j, 4 + 4 * j
        put(11, c1, 11, c2, h, bold=True, font_size=6, align="center", valign="vcenter",
            text_wrap=True)
        for k in range(11, 16):
            cv.box(R(k), C(c1), R(k), C(c2), THIN)
    cv.box(R(11), C(1), R(15), C(24), MEDIUM)

    # --- FECHA DE CALIDAD / FIRMA ---------------------------------------
    put(11, 26, 11, 47, "  FECHA DE CALIDAD / FIRMA", font_size=7, **black)
    for k in range(12, 15):
        cv.hline(R(k), C(26), C(47))
    cv.box(R(11), C(26), R(15), C(47), MEDIUM)

    # --- TEST DATA -------------------------------------------------------
    put(17, 1, 17, 11, " TEST DATA", font_size=9, **black)
    put(17, 12, 17, 24, "ÚLTIMA PIEZA", font_size=7, align="center", **black)
    tests = ["VOLTS", "AMPS", "MEDIUM", "PRESSURE MIN (PSI)", "PRESSURE MAX (PSI)",
             "EXTERNAL LEAK", "TEST DATE", "SNOOP"]
    for j, t in enumerate(tests):
        k = 18 + j
        put(k, 1, k, 11, " " + t, **small)
        cv.hline(R(k), C(1), C(24))
        cv.vline(C(12), R(k), R(k))
    cv.box(R(17), C(1), R(25), C(24), MEDIUM)

    # --- Cuadro QA -------------------------------------------------------
    def checkbox(k, c):
        cv.box(R(k), C(c), R(k), C(c), THIN)

    put(17, 26, 17, 33, " COMPLETE Y/N", bold=True, font_size=6, valign="vcenter")
    checkbox(17, 34)
    put(17, 35, 17, 35, "Y", font_size=6, align="center", valign="vcenter")
    checkbox(17, 37)
    put(17, 38, 17, 38, "N", font_size=6, align="center", valign="vcenter")
    put(17, 44, 17, 46, "QA", bold=True, font_size=7, align="right", valign="vcenter")
    put(19, 26, 19, 33, " # EMP EMPAQUE", bold=True, font_size=6, valign="bottom")
    cv.hline(R(19), C(34), C(46))
    put(21, 26, 21, 29, " NOTAS", bold=True, font_size=6, valign="bottom")
    cv.hline(R(21), C(30), C(46))
    put(23, 26, 23, 29, " # LOC", bold=True, font_size=6, valign="bottom")
    cv.hline(R(23), C(30), C(46))
    put(25, 26, 25, 32, " CERTIFICADO", bold=True, font_size=6, valign="vcenter")
    checkbox(25, 33)
    put(25, 34, 25, 37, "APLICA", font_size=6, valign="vcenter")
    checkbox(25, 39)
    put(25, 40, 25, 42, "N/A", font_size=6, valign="vcenter")
    cv.box(R(17), C(26), R(25), C(47), MEDIUM)

    put(26, 26, 26, 47, "Falla → HOLD · FME-8.5.1-XX Rev 0 + FME-8.5.1-38 Rev 1",
        font_size=5, valign="vcenter")


def build():
    wb = xlsxwriter.Workbook(OUT)
    wb.set_vba_name("ThisWorkbook")
    wb.set_properties({"title": "Etiqueta Orden Operador - Code 128",
                       "subject": "Impresion de etiquetas 4x6 con codigo de barras Code 128"})

    et = wb.add_worksheet("ETIQUETA")
    li = wb.add_worksheet("LISTA")
    cf = wb.add_worksheet("CONFIG")
    et.set_vba_name("shEtiqueta")
    li.set_vba_name("shLista")
    cf.set_vba_name("shConfig")

    mods = [vba_project.load_bas(os.path.join(ROOT, "vba", f))
            for f in ("ModCode128.bas", "ModImpresion.bas")]
    with open(os.path.join(ROOT, "vba", "ThisWorkbook.cls"), encoding="utf-8") as f:
        wb_code = f.read()
    bin_path = OUT + ".vbaProject.bin"
    with open(bin_path, "wb") as f:
        f.write(vba_project.build_vba_project(mods, "ThisWorkbook",
                                              ["shEtiqueta", "shLista", "shConfig"], wb_code))
    wb.add_vba_project(bin_path)

    names = {}

    # ===================== ETIQUETA ======================================
    # Etiqueta girada a la izquierda (columnas B..AB), panel de captura a la derecha.
    et.hide_gridlines(2)
    et.set_column_pixels(0, 0, 16)
    for k, h in enumerate(LABEL_ROWS):
        et.set_column_pixels(LCOL0 + k, LCOL0 + k, round(h * 4 / 3))
    gap = LCOL0 + len(LABEL_ROWS)
    et.set_column_pixels(gap, gap, 24)
    P = gap + 1                       # columnas del panel: CAMPO, DATO, PRUEBA, ESTADO
    for j, w in enumerate([112, 190, 150, 130]):
        et.set_column_pixels(P + j, P + j, w)
    et.set_row(0, 26)
    for i in range(LABEL_COLS):
        et.set_row(LROW0 + i, ROW_PT)

    cv = Canvas(wb, et)
    pf = dict(font_name="Calibri", font_size=10, valign="vcenter")

    def prow(i):
        return LROW0 + i

    cv.put(0, 0, 0, P + 3, "  ORDEN OPERADOR  ·  Etiqueta 4 x 6 con Code 128  ·  v2 (vertical)",
           font_name="Calibri", font_size=13, bold=True, font_color="#FFFFFF",
           bg_color="#1F3864", valign="vcenter")
    for j, h in enumerate(["CAMPO", "DATO (escribe o escanea)", "PRUEBA: escanea aquí",
                           "ESTADO"]):
        cv.put(prow(0), P + j, prow(1), P + j, h, bold=True, bg_color="#D9D9D9",
               align="center", border=THIN, **pf)

    fields = [
        ("W.O. / JOB", "In_JOB", "Scan_JOB", "BC_JOB", "St_JOB"),
        ("OB", "In_OB", "Scan_OB", "BC_OB", "St_OB"),
        ("QTY TOTAL", "In_QTY", "Scan_QTY", "BC_QTY", "St_QTY"),
        ("P/N (MODELO)", "In_PN", "Scan_PN", "BC_PN", "St_PN"),
        ("HEAT CODE", "In_HEATCODE", None, None, None),
        ("HEAT LOT", "In_HEATLOT", None, None, None),
        ("COPIAS", "In_COPIAS", None, None, None),
    ]
    inp = dict(bg_color="#FFF2CC", border=THIN, locked=False, num_format="@", bold=True,
               font_name="Calibri", font_size=12, valign="vcenter")
    scan = dict(bg_color="#DDEBF7", border=THIN, locked=False, num_format="@",
                font_name="Calibri", font_size=11, valign="vcenter")
    for j, (title, name, scan_name, bc_name, st_name) in enumerate(fields):
        r1, r2 = prow(2 + 3 * j), prow(4 + 3 * j)
        cv.put(r1, P, r2, P, " " + title, bold=True, border=THIN, **pf)
        if name == "In_COPIAS":
            cv.put(r1, P + 1, r2, P + 1, 1, **dict(inp, num_format="0", align="left"))
        else:
            cv.put(r1, P + 1, r2, P + 1, None, **inp)
        names[name] = (r1, P + 1)
        if scan_name:
            cv.put(r1, P + 2, r2, P + 2, None, **scan)
            names[scan_name] = (r1, P + 2)
            f = ('=IF(TRIM({i})="","",IF(ISERROR(CODE128("0")),"HABILITA MACROS",'
                 'IF({b}="","CARÁCTER NO VÁLIDO",IF(TRIM({s})="","listo",'
                 'IF(UPPER(TRIM({s}))=UPPER(TRIM({i})),"OK","NO COINCIDE")))))'
                 ).format(i=name, s=scan_name, b=bc_name)
            cv.put(r1, P + 3, r2, P + 3, f, formula=True, cached="", bold=True,
                   align="center", border=THIN, **dict(pf, font_size=9))
            names[st_name] = (r1, P + 3)
        else:
            note = {"In_HEATCODE": "opcional: vacío = línea para escribir",
                    "In_HEATLOT": "opcional: vacío = línea para escribir",
                    "In_COPIAS": "copias de cada etiqueta (1 a 500)"}[name]
            cv.put(r1, P + 2, r2, P + 3, " " + note, italic=True, font_color="#7F7F7F",
                   border=THIN, **pf)

    design = Canvas(None, None)
    design.names = {}
    label_layout(design)
    names.update(turn_into(design, cv))
    cv.flush()

    green = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
    red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    status_rng = "%s:%s" % (xlsxwriter.utility.xl_rowcol_to_cell(prow(2), P + 3),
                            xlsxwriter.utility.xl_rowcol_to_cell(prow(13), P + 3))
    et.conditional_format(status_rng, {"type": "text", "criteria": "begins with",
                                       "value": "OK", "format": green})
    for word in ("NO ", "CAR", "HAB"):
        et.conditional_format(status_rng, {"type": "text", "criteria": "begins with",
                                           "value": word, "format": red})
    r, c = names["In_COPIAS"]
    et.data_validation(r, c, r, c, {
        "validate": "integer", "criteria": "between", "minimum": 1, "maximum": 500,
        "error_title": "Copias", "error_message": "Escribe un número de 1 a 500."})

    et.insert_button(prow(24), P + 1, {"macro": "Imprimir_Etiqueta", "caption": "IMPRIMIR",
                                       "width": 182, "height": 52, "x_offset": 4})
    et.insert_button(prow(24), P + 2, {"macro": "Vista_Previa", "caption": "Vista previa",
                                       "width": 142, "height": 24, "x_offset": 4})
    et.insert_button(prow(26), P + 2, {"macro": "Limpiar_Etiqueta", "caption": "Limpiar",
                                       "width": 142, "height": 24, "x_offset": 4,
                                       "y_offset": 4})
    et.insert_button(prow(24), P + 3, {"macro": "Elegir_Impresora",
                                       "caption": "Elegir impresora",
                                       "width": 122, "height": 24, "x_offset": 4})
    help_fmt = wb.add_format({"font_name": "Calibri", "font_size": 10, "text_wrap": True,
                              "valign": "top", "font_color": "#404040"})
    et.merge_range(prow(31), P, prow(46), P + 3,
                   "PASOS\n"
                   "1. Captura o escanea W.O./JOB, OB, QTY TOTAL y P/N.\n"
                   "2. Revisa la etiqueta de la izquierda: así sale impresa, en la misma "
                   "posición que Label Matrix.\n"
                   "3. Pon las COPIAS y presiona IMPRIMIR (se imprime vertical, 4 x 6 in).\n"
                   "4. Escanea la etiqueta impresa en la columna PRUEBA: debe decir OK.\n\n"
                   "Código: Code 128 Auto, dígito verificador mod 103.\n"
                   "Varias órdenes: hoja LISTA.   Fuente: hoja CONFIG.",
                   help_fmt)

    first = xlsxwriter.utility.xl_rowcol_to_cell(LROW0, LCOL0)
    last = xlsxwriter.utility.xl_rowcol_to_cell(LROW0 + LABEL_COLS - 1,
                                                LCOL0 + len(LABEL_ROWS) - 1)
    et.print_area("%s:%s" % (first, last))
    et.set_portrait()
    et.set_margins(left=0, right=0, top=0, bottom=0)
    et.set_header("", {"margin": 0})
    et.set_footer("", {"margin": 0})
    et.center_horizontally()
    et.center_vertically()
    et.fit_to_pages(1, 1)
    et.protect("")
    et.activate()
    r, c = names["In_JOB"]
    et.set_selection(r, c, r, c)

    # ===================== LISTA =========================================
    li.set_column(0, 0, 5)
    li.set_column(1, 2, 14)
    li.set_column(3, 3, 11)
    li.set_column(4, 4, 26)
    li.set_column(5, 6, 14)
    li.set_column(7, 7, 8)
    li.set_column(8, 8, 26)
    title = wb.add_format({"bold": True, "font_size": 13, "font_color": "#FFFFFF",
                           "bg_color": "#1F3864", "valign": "vcenter"})
    note = wb.add_format({"italic": True, "font_color": "#595959", "text_wrap": True,
                          "valign": "top"})
    head = wb.add_format({"bold": True, "bg_color": "#D9D9D9", "border": THIN,
                          "align": "center", "valign": "vcenter", "text_wrap": True})
    txt = wb.add_format({"num_format": "@", "border": THIN, "border_color": "#BFBFBF"})
    num = wb.add_format({"border": THIN, "border_color": "#BFBFBF", "align": "center"})
    idx = wb.add_format({"font_color": "#7F7F7F", "align": "center"})
    li.set_row(0, 24)
    li.merge_range(0, 0, 0, 8, "  LISTA DE IMPRESIÓN — varias órdenes", title)
    li.set_row(1, 30)
    li.merge_range(1, 0, 1, 8,
                   "Captura o pega una orden por fila (desde la fila 5). IMPRIMIR LISTA "
                   "imprime las filas cuyo ESTADO no dice IMPRESO. Para reimprimir una "
                   "fila, borra su ESTADO.", note)
    li.set_row(3, 30)
    for c, h in enumerate(["#", "W.O. / JOB", "OB", "QTY TOTAL", "P/N (MODELO)",
                           "HEAT CODE", "HEAT LOT", "COPIAS", "ESTADO"]):
        li.write(3, c, h, head)
    for r in range(4, 504):
        li.write_formula(r, 0, '=IF(B%d="","",ROW()-4)' % (r + 1), idx, "")
        for c in range(1, 7):
            li.write_blank(r, c, None, txt)
        li.write_blank(r, 7, None, num)
        li.write_blank(r, 8, None, txt)
    li.data_validation(4, 7, 503, 7, {
        "validate": "integer", "criteria": "between", "minimum": 1, "maximum": 500,
        "ignore_blank": True, "error_title": "Copias",
        "error_message": "Escribe un número de 1 a 500 (vacío = 1)."})
    li.freeze_panes(4, 0)
    li.set_landscape()
    li.fit_to_pages(1, 0)
    li.repeat_rows(3)
    li.insert_button(3, 10, {"macro": "Imprimir_Lista", "caption": "IMPRIMIR LISTA",
                             "width": 150, "height": 34})
    li.insert_button(6, 10, {"macro": "Limpiar_Lista", "caption": "Limpiar lista",
                             "width": 150, "height": 24})

    # ===================== CONFIG ========================================
    cf.set_column(0, 0, 2)
    cf.set_column(1, 1, 34)
    cf.set_column(2, 2, 46)
    cf.set_column(3, 3, 3)
    cf.set_column(4, 8, 12)
    lab = wb.add_format({"bold": True, "valign": "vcenter", "border": THIN})
    val = wb.add_format({"bg_color": "#FFF2CC", "border": THIN, "valign": "vcenter",
                         "bold": True})
    valnum = wb.add_format({"bg_color": "#FFF2CC", "border": THIN, "valign": "vcenter",
                            "bold": True, "align": "left", "num_format": "0.0"})
    info = wb.add_format({"border": THIN, "valign": "vcenter", "text_wrap": True})
    sec = wb.add_format({"bold": True, "font_color": "#1F3864", "font_size": 12,
                         "bottom": MEDIUM, "bottom_color": "#1F3864"})
    bcf = wb.add_format({"font_name": BARCODE_FONT, "font_size": BARCODE_SIZE,
                         "valign": "top", "border": THIN, "shrink": True})
    wrap = wb.add_format({"text_wrap": True, "valign": "top"})

    cf.set_row(0, 24)
    cf.merge_range(0, 0, 0, 8, "  CONFIGURACIÓN DEL CÓDIGO DE BARRAS", title)
    cf.write(2, 1, "FUENTE", sec)
    cf.write(2, 2, "", sec)
    rows = [
        ("Fuente de código de barras", "cfgFuente", BARCODE_FONT, val),
        ("Mapeo de caracteres de la fuente", "cfgMapeo", MAPEO_DEFAULT, val),
        ("Tamaño de la fuente (pt)", "cfgTamano", BARCODE_SIZE, valnum),
    ]
    for j, (label, name, value, f) in enumerate(rows):
        cf.set_row(3 + j, 20)
        cf.write(3 + j, 1, label, lab)
        cf.write(3 + j, 2, value, f)
        names[name] = (3 + j, 2, "CONFIG")
    cf.data_validation(4, 2, 4, 2, {"validate": "list", "source": MAPEOS})
    cf.fit_to_pages(1, 0)
    cf.set_row(6, 20)
    cf.write(6, 1, "Simbología", lab)
    cf.write(6, 2, "Code 128 Auto (B/C), verificador mód. 103", info)
    cf.insert_button(3, 4, {"macro": "Aplicar_Fuente", "caption": "Aplicar fuente",
                            "width": 150, "height": 30})

    cf.write(8, 1, "PRUEBA", sec)
    cf.write(8, 2, "", sec)
    sample = "50426529"
    cf.set_row(9, 20)
    cf.write(9, 1, "Texto de prueba", lab)
    cf.write_string(9, 2, sample, wb.add_format({"bg_color": "#FFF2CC", "border": THIN,
                                                 "bold": True, "num_format": "@"}))
    names["cfgPruebaTexto"] = (9, 2, "CONFIG")
    cf.set_row(10, 20)
    cf.write(10, 1, "Texto codificado (lo que recibe la fuente)", lab)
    cf.write_formula(10, 2, '=IFERROR(CODE128(cfgPruebaTexto,cfgMapeo),"HABILITA MACROS")',
                     info, code128.encode(sample, MAPEO_DEFAULT))
    cf.set_row(11, 20)
    cf.write(11, 1, "Dígito verificador (mód. 103)", lab)
    cf.write_formula(11, 2, '=IFERROR(CODE128_CHECK(cfgPruebaTexto),"")',
                     wb.add_format({"border": THIN, "align": "left", "valign": "vcenter"}),
                     code128.check_digit(sample))
    cf.set_row(12, 30)
    cf.write(12, 1, "Código de barras", lab)
    cf.write_formula(12, 2, '=IFERROR(CODE128(cfgPruebaTexto,cfgMapeo),"")', bcf,
                     code128.encode(sample, MAPEO_DEFAULT))
    names["cfgPruebaBarra"] = (12, 2, "CONFIG")
    cf.set_row(13, 45)
    cf.merge_range(13, 1, 13, 2,
                   "Si en \"Código de barras\" ves letras en lugar de barras, la fuente no está "
                   "instalada o el nombre en \"Fuente\" no coincide: escribe el nombre exacto como "
                   "aparece en la lista de fuentes de Excel y presiona Aplicar fuente.", wrap)

    cf.write(15, 1, "INSTRUCCIONES", sec)
    cf.write(15, 2, "", sec)
    steps = [
        "1. Si Excel dice que bloqueó las macros: cierra el archivo, clic derecho > "
        "Propiedades > marca \"Desbloquear\" > Aceptar, y ábrelo de nuevo. Luego "
        "\"Habilitar contenido\".",
        "2. La fuente debe estar instalada en cada PC que imprima. Por defecto se usa "
        "\"Libre Barcode 128\" (mapeo LIBRE / IDAUTOMATION). Si tu fuente es la "
        "\"Code 128\" de grandzebu, elige el mapeo GRANDZEBU.",
        "3. Impresora Zebra ZT230 (203 dpi): en Preferencias de impresión pon el tamaño "
        "de etiqueta 4.00 x 6.00 in y orientación Vertical. La hoja ETIQUETA ya trae la "
        "etiqueta girada como en Label Matrix y se imprime vertical, márgenes 0, ajustada "
        "a 1 página.",
        "4. Con 35.5 pt cada barra delgada mide 3 puntos de la Zebra (0.015 in). No uses "
        "menos de 24 pt; un P/N muy largo se reduce solo para caber.",
        "5. Code 128 solo acepta letras sin acento, números y símbolos ASCII (sin Ñ). "
        "El ESTADO avisa si un dato no se puede codificar.",
        "6. Después de imprimir, escanea cada código en la columna PRUEBA de ETIQUETA: "
        "debe decir OK. Así se confirma el código y el dígito verificador.",
    ]
    for j, s in enumerate(steps):
        cf.set_row(16 + j, 46)
        cf.merge_range(16 + j, 1, 16 + j, 2, s, wrap)

    # ===================== nombres =======================================
    for name, ref in names.items():
        sheet = ref[2] if len(ref) > 2 else "ETIQUETA"
        cell = xlsxwriter.utility.xl_rowcol_to_cell(ref[0], ref[1], row_abs=True, col_abs=True)
        wb.define_name(name, "=%s!%s" % (sheet, cell))

    wb.close()
    os.remove(bin_path)
    print("ok:", OUT)


if __name__ == "__main__":
    build()

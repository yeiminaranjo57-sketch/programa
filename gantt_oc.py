"""
GANTT DE PROGRAMACIÓN POR TALLER - Órdenes de corte (OC)
========================================================

Qué hace este archivo (en palabras sencillas):
  1. Lee el archivo de entrada (por defecto "entradas/PRUEBA_GANTT.xlsx"):
        - Hoja "LISTA OC":  una fila por OC y talle.
        - Hoja "CAPACIDAD": por MODELO + COLOR, operadores, tiempo y TALLER.
  2. Suma la CANTIDAD PLANIFICADA de cada OC (todos los talles juntos).
  3. Calcula la capacidad diaria = OPERADORES x 495 minutos / TIEMPO
     (la misma fórmula de la hoja CAPACIDAD).
  4. Cada TALLER tiene una o varias LÍNEAS de producción (hoja "TALLERES"
     del archivo de entrada, o la tabla TALLERES_POR_DEFECTO de abajo).
     La capacidad de la hoja CAPACIDAD es la de UNA línea. Cada OC puede
     ir como máximo en N líneas a la vez (CASEROS: 7 líneas, máx. 2 por OC).
  5. Las OC se toman en el orden de la hoja LISTA OC. A cada una se le dan
     las líneas que se liberan primero y la cantidad se reparte para que
     terminen a la vez. Solo se usa una segunda línea si la OC necesita más
     de un día de una línea. Si una OC termina a mitad de día, la siguiente
     usa el resto de ese día en esa línea.
  6. Solo cuenta días hábiles: sin sábados, domingos ni feriados.
     Los feriados se toman de la hoja "FERIADOS" del archivo de entrada si
     existe (columna FECHA); si no, se usa la lista de abajo (Argentina).
  7. Guarda "resultados/Programa_Gantt_OC.xlsx" con el Gantt por línea y
     por OC, el programa por OC, el detalle diario y la carga por línea.

Cómo se usa:
  python gantt_oc.py                       (usa el archivo y fecha por defecto)
  python gantt_oc.py otro_archivo.xlsx     (otro archivo de entrada)
  python gantt_oc.py otro_archivo.xlsx 2026-10-19   (otra fecha de inicio)
"""

import math
import os
import sys
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_ENTRADA = os.path.join(CARPETA, "entradas", "PRUEBA_GANTT.xlsx")
ARCHIVO_SALIDA = os.path.join(CARPETA, "resultados", "Programa_Gantt_OC.xlsx")
FECHA_INICIO = date(2026, 10, 9)
MINUTOS_POR_DIA = 495

# Feriados nacionales de Argentina (se pueden reemplazar con una hoja FERIADOS
# en el archivo de entrada). Los días "no laborables" (7/12/2026, jueves santo)
# NO están: en esos días decide la empresa.
FERIADOS_POR_DEFECTO = {
    "2026-10-12": "Día del Respeto a la Diversidad Cultural",
    "2026-11-09": "Feriado por visita del Papa (Decreto 1103/2026) - verificar",
    "2026-11-23": "Día de la Soberanía Nacional (trasladado del 20/11)",
    "2026-12-08": "Inmaculada Concepción de María",
    "2026-12-25": "Navidad",
    "2027-01-01": "Año Nuevo",
    "2027-02-08": "Carnaval",
    "2027-02-09": "Carnaval",
    "2027-03-24": "Día Nacional de la Memoria por la Verdad y la Justicia",
    "2027-03-26": "Viernes Santo",
    "2027-04-02": "Día del Veterano y de los Caídos en Malvinas",
    "2027-05-25": "Revolución de Mayo",
    "2027-06-20": "Paso a la Inmortalidad del Gral. Belgrano",
    "2027-07-09": "Día de la Independencia",
    "2027-12-08": "Inmaculada Concepción de María",
    "2027-12-25": "Navidad",
}

# Líneas de cada taller: (cantidad de líneas, máximo de líneas por OC).
# Se puede reemplazar con una hoja TALLERES en el archivo de entrada con las
# columnas TALLER, LINEAS y MAX_LINEAS_POR_OC. Un taller que no esté aquí
# trabaja con 1 línea.
TALLERES_POR_DEFECTO = {
    "CASEROS": (7, 2),
    "OLIDEN": (1, 1),
}

# Colores de las barras (uno por taller) y de alerta
COLORES_OC = ["4F81BD", "9BBB59", "F79646", "8064A2", "4BACC6", "2C4D75", "77933C", "B65708", "604A7B", "31859C"]
COLOR_ATRASO = "FF0000"
COLOR_ENCABEZADO = "1F3864"
COLOR_TALLER_FILA = "D9E1F2"


def texto(valor):
    return "" if pd.isna(valor) else str(valor).strip()


def a_fecha(valor):
    if pd.isna(valor) or valor == "":
        return None
    if isinstance(valor, (datetime, pd.Timestamp)):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return datetime.strptime(str(valor).strip()[:10], "%d/%m/%Y").date()


# ---------------------------------------------------------------- lectura

def leer_capacidad(archivo):
    cap = pd.read_excel(archivo, sheet_name="CAPACIDAD")
    cap.columns = [texto(c).upper() for c in cap.columns]
    cap["MODELO"] = cap["MODELO"].map(texto)
    cap["COLOR"] = cap["COLOR"].map(texto)
    cap["TALLER"] = cap["TALLER"].map(texto)
    # La columna CAPACIDAD es una fórmula; se recalcula para no depender de Excel.
    cap["CAP_DIA"] = cap["OPERADORES"].astype(float) * MINUTOS_POR_DIA / cap["TIEMPO"].astype(float)
    return cap


def leer_feriados(archivo):
    try:
        hoja = pd.read_excel(archivo, sheet_name="FERIADOS")
        hoja.columns = [texto(c).upper() for c in hoja.columns]
        feriados = {}
        for _, fila in hoja.iterrows():
            f = a_fecha(fila["FECHA"])
            if f:
                feriados[f] = texto(fila.get("DESCRIPCION", "")) or "Feriado"
        return feriados, "hoja FERIADOS del archivo de entrada"
    except ValueError:
        return {date.fromisoformat(k): v for k, v in FERIADOS_POR_DEFECTO.items()}, "lista de Argentina incluida en el programa"


def leer_talleres(archivo):
    try:
        hoja = pd.read_excel(archivo, sheet_name="TALLERES")
    except ValueError:
        return dict(TALLERES_POR_DEFECTO), "tabla incluida en el programa"
    hoja.columns = [texto(c).upper() for c in hoja.columns]
    talleres = {}
    for _, fila in hoja.iterrows():
        if texto(fila["TALLER"]):
            lineas = int(fila["LINEAS"])
            talleres[texto(fila["TALLER"])] = (lineas, min(lineas, int(fila.get("MAX_LINEAS_POR_OC", lineas))))
    return talleres, "hoja TALLERES del archivo de entrada"


def leer_ocs(archivo, cap, advertencias):
    lista = pd.read_excel(archivo, sheet_name="LISTA OC")
    lista.columns = [texto(c).upper() for c in lista.columns]
    lista = lista[lista["OC"].notna()].copy()
    lista["MODELO"] = lista["MODELO"].map(texto)
    lista["COLOR"] = lista["COLOR"].map(texto)

    ocs = []
    for oc, g in lista.groupby("OC", sort=False):  # respeta el orden de la hoja
        primera = g.iloc[0]
        if g["MODELO"].nunique() > 1 or g["COLOR"].nunique() > 1:
            advertencias.append(f"OC {oc}: tiene más de un modelo/color; se usa {primera['MODELO']} {primera['COLOR']}.")
        fila_cap = cap[(cap["MODELO"] == primera["MODELO"]) & (cap["COLOR"] == primera["COLOR"])]
        if fila_cap.empty:
            fila_cap = cap[cap["MODELO"] == primera["MODELO"]]
            if not fila_cap.empty:
                advertencias.append(f"OC {oc}: color {primera['COLOR']} no está en CAPACIDAD; se usa la del modelo.")
        if fila_cap.empty:
            advertencias.append(f"OC {int(oc)}: modelo {primera['MODELO']} sin capacidad ni taller; NO se programa.")
            continue
        fila_cap = fila_cap.iloc[0]
        cantidad = float(g["CANTIDAD PLANIFICADA"].fillna(0).sum())
        if cantidad <= 0:
            advertencias.append(f"OC {int(oc)}: cantidad planificada 0; NO se programa.")
            continue
        ocs.append({
            "OC": int(oc),
            "TALLER": fila_cap["TALLER"],
            "CLIENTE": texto(primera.get("CLIENTE")).replace("\xa0", " "),
            "DESCRIPCION": texto(primera.get("DESCRIPCION")),
            "FAMILIA": texto(primera.get("FAMILIA")),
            "MODELO": primera["MODELO"],
            "COLOR": primera["COLOR"],
            "ESTADO OC": texto(primera.get("ESTADO OC")),
            "TALLES": len(g),
            "CANTIDAD": int(round(cantidad)),
            "CAP_DIA": float(fila_cap["CAP_DIA"]),
            "FECHA FIN OC": a_fecha(primera.get("FECHA FIN")),
        })
    return ocs


# ---------------------------------------------------------------- cálculo

def dias_habiles(desde, feriados):
    """Generador infinito de días hábiles (lunes a viernes, sin feriados)."""
    d = desde
    while True:
        if d.weekday() < 5 and d not in feriados:
            yield d
        d += timedelta(days=1)


def repartir_por_dia(inicio, cantidad, cap_dia):
    """Unidades por día de un tramo que empieza en 'inicio' (días hábiles con
    decimales: 2,5 = mitad del tercer día). Redondeo acumulado para que la
    suma sea exactamente la cantidad."""
    fin = inicio + cantidad / cap_dia
    por_dia, hecho, k = {}, 0, int(math.floor(inicio))
    while k < fin - 1e-9:
        hasta = min(fin, k + 1)
        acumulado = round((hasta - inicio) * cap_dia) if hasta < fin else cantidad
        if acumulado - hecho > 0:
            por_dia[k] = acumulado - hecho
        hecho = acumulado
        k += 1
    return fin, por_dia


def nombre_linea(i):
    return f"L{i + 1}"


def programar(ocs, fecha_inicio, feriados, talleres):
    """Asigna cada OC (en orden) a las líneas que se liberan primero, hasta el
    máximo permitido, repartiendo la cantidad para que terminen a la vez."""
    libre = {}  # taller -> [momento en que se libera cada línea]
    for oc in ocs:
        n_lineas, max_por_oc = talleres.get(oc["TALLER"], (1, 1))
        libres = libre.setdefault(oc["TALLER"], [0.0] * n_lineas)
        orden = sorted(range(n_lineas), key=lambda i: (libres[i], i))
        q, r = oc["CANTIDAD"], oc["CAP_DIA"]

        elegidas, fin = orden[:1], libres[orden[0]] + q / r
        if q > r:  # una OC de menos de un día de una línea no se divide
            for k in range(2, max_por_oc + 1):
                cand = orden[:k]
                fin_k = (sum(libres[i] for i in cand) + q / r) / k
                if fin_k <= libres[cand[-1]] + 1e-9:
                    break  # la línea extra se libera después: no ayuda
                elegidas, fin = cand, fin_k

        # Unidades enteras por línea; la última línea se lleva el resto
        cantidades, resto = [], q
        for j, i in enumerate(elegidas):
            u = resto if j == len(elegidas) - 1 else min(resto, int(round(r * (fin - libres[i]))))
            cantidades.append(u)
            resto -= u

        oc["TRAMOS"], oc["POR_DIA"] = [], {}
        for i, u in zip(elegidas, cantidades):
            if u <= 0:
                continue
            inicio = libres[i]
            fin_tramo, por_dia = repartir_por_dia(inicio, u, r)
            libres[i] = fin_tramo
            oc["TRAMOS"].append({"LINEA": i, "INICIO": inicio, "FIN": fin_tramo, "UNIDADES": u, "POR_DIA": por_dia})
            for k, v in por_dia.items():
                oc["POR_DIA"][k] = oc["POR_DIA"].get(k, 0) + v
        oc["LINEAS"] = "+".join(nombre_linea(t["LINEA"]) for t in sorted(oc["TRAMOS"], key=lambda t: t["LINEA"]))
        oc["DIAS"] = max(t["FIN"] for t in oc["TRAMOS"]) - min(t["INICIO"] for t in oc["TRAMOS"])

    total_dias = int(math.ceil(max(max(v) for v in libre.values()))) if libre else 0
    gen = dias_habiles(fecha_inicio, feriados)
    calendario = [next(gen) for _ in range(total_dias)]

    for oc in ocs:
        dias = sorted(oc["POR_DIA"])
        oc["INICIO PROG"] = calendario[dias[0]]
        oc["FIN PROG"] = calendario[dias[-1]]
        limite = oc["FECHA FIN OC"]
        if limite is None:
            oc["ESTADO"], oc["ATRASO"] = "SIN FECHA FIN", None
        elif oc["FIN PROG"] <= limite:
            oc["ESTADO"], oc["ATRASO"] = "A TIEMPO", 0
        else:
            oc["ESTADO"] = "ATRASADA"
            oc["ATRASO"] = sum(1 for d in calendario if limite < d <= oc["FIN PROG"])
    return calendario


# ---------------------------------------------------------------- salida

FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
DIAS_SEM = ["LU", "MA", "MI", "JU", "VI"]
FMT_FECHA = "dd/mm/yyyy"


def estilo_encabezado(celda):
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor=COLOR_ENCABEZADO)
    celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def tabla(ws, columnas, filas, anchos=None, formatos=None):
    for j, nombre in enumerate(columnas, 1):
        estilo_encabezado(ws.cell(row=1, column=j, value=nombre))
    for i, fila in enumerate(filas, 2):
        for j, valor in enumerate(fila, 1):
            c = ws.cell(row=i, column=j, value=valor)
            if formatos and columnas[j - 1] in formatos:
                c.number_format = formatos[columnas[j - 1]]
    for j, nombre in enumerate(columnas, 1):
        ws.column_dimensions[get_column_letter(j)].width = (anchos or {}).get(nombre, max(10, len(nombre) + 2))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def encabezado_calendario(ws, fijas, calendario):
    """Filas 1-3: columnas fijas + mes / día de la semana / fecha."""
    c0 = len(fijas) + 1
    for j, nombre in enumerate(fijas, 1):
        ws.merge_cells(start_row=1, start_column=j, end_row=3, end_column=j)
        estilo_encabezado(ws.cell(row=1, column=j, value=nombre))
    inicio_mes = c0
    for k, d in enumerate(calendario):
        col = c0 + k
        estilo_encabezado(ws.cell(row=2, column=col, value=DIAS_SEM[d.weekday()]))
        c = ws.cell(row=3, column=col, value=d)
        estilo_encabezado(c)
        c.number_format = "dd/mm"
        c.alignment = Alignment(horizontal="center", text_rotation=90)
        if k == len(calendario) - 1 or calendario[k + 1].month != d.month:
            if col > inicio_mes:
                ws.merge_cells(start_row=1, start_column=inicio_mes, end_row=1, end_column=col)
            estilo_encabezado(ws.cell(row=1, column=inicio_mes, value=f"{MESES[d.month - 1]} {d.year}"))
            inicio_mes = col + 1
        ws.column_dimensions[get_column_letter(col)].width = 4.2
    ws.row_dimensions[3].height = 42
    ws.freeze_panes = ws.cell(row=4, column=c0)
    return c0


def celda_barra(c, valor, color):
    c.value = valor
    c.fill = PatternFill("solid", fgColor=color)
    c.font = Font(color="FFFFFF", size=8, bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90, wrap_text=True)


def fila_taller(ws, fila, c0, calendario, taller, del_taller, n_lineas, totales):
    valores = {1: taller, 2: f"{n_lineas} línea(s)", 3: f"{len(del_taller)} OC",
               4: sum(o["CANTIDAD"] for o in del_taller)}
    for col in range(1, c0 + len(calendario)):
        c = ws.cell(row=fila, column=col, value=valores.get(col))
        c.font = Font(bold=True, size=8 if col >= c0 else 11)
        c.fill = PatternFill("solid", fgColor=COLOR_TALLER_FILA)
        c.border = BORDE
        if col >= c0 and (col - c0) in totales:
            c.value = totales[col - c0]
            c.alignment = Alignment(horizontal="center", text_rotation=90)


def hoja_gantt_lineas(wb, ocs, calendario, talleres_orden, talleres):
    """Gantt por TALLER: una fila por línea con la OC que trabaja cada día,
    y debajo las unidades de ese día."""
    ws = wb.active
    ws.title = "GANTT_LINEAS"
    c0 = encabezado_calendario(ws, ["TALLER", "LÍNEA", "OC", "UNIDADES"], calendario)
    fila = 4
    for taller in talleres_orden:
        del_taller = [o for o in ocs if o["TALLER"] == taller]
        n_lineas = talleres.get(taller, (1, 1))[0]
        totales = {}
        for o in del_taller:
            for k, u in o["POR_DIA"].items():
                totales[k] = totales.get(k, 0) + u
        fila_taller(ws, fila, c0, calendario, taller, del_taller, n_lineas, totales)
        fila += 1
        for i in range(n_lineas):
            tramos = [(n, o, t) for n, o in enumerate(del_taller) for t in o["TRAMOS"] if t["LINEA"] == i]
            dia = {}
            for n, o, t in tramos:
                for k, u in t["POR_DIA"].items():
                    dia.setdefault(k, []).append((n, o, u))
            ws.cell(row=fila, column=1, value=taller)
            ws.cell(row=fila, column=2, value=nombre_linea(i)).font = Font(bold=True)
            ws.cell(row=fila, column=3, value=len(tramos))
            ws.cell(row=fila, column=4, value=sum(t["UNIDADES"] for _, _, t in tramos))
            ws.cell(row=fila + 1, column=2, value="unidades").font = Font(italic=True, size=8)
            for k, lista in dia.items():
                n, o, _ = lista[-1]
                tarde = any(x[1]["FECHA FIN OC"] and calendario[k] > x[1]["FECHA FIN OC"] for x in lista)
                color = COLOR_ATRASO if tarde else COLORES_OC[n % len(COLORES_OC)]
                celda_barra(ws.cell(row=fila, column=c0 + k), " / ".join(str(x[1]["OC"]) for x in lista), color)
                u = ws.cell(row=fila + 1, column=c0 + k, value=sum(x[2] for x in lista))
                u.font = Font(size=7)
                u.alignment = Alignment(horizontal="center", text_rotation=90)
            for col in range(1, c0 + len(calendario)):
                ws.cell(row=fila, column=col).border = BORDE
                if col < c0:
                    ws.cell(row=fila, column=col).alignment = Alignment(vertical="center")
            ws.row_dimensions[fila].height = 70
            fila += 2
        fila += 1
    for j, a in enumerate([11, 9, 6, 10], 1):
        ws.column_dimensions[get_column_letter(j)].width = a
    notas = ["Cada celda de una línea muestra la OC que trabaja ese día (dos OC si una termina y otra empieza).",
             "El mismo color = la misma OC (aunque esté en dos líneas). ROJO = días después de la FECHA FIN de la OC.",
             "Fila azul clara = total del taller por día. Solo días hábiles (sin sábados, domingos ni feriados)."]
    for j, n in enumerate(notas):
        ws.cell(row=fila + 1 + j, column=1, value=n)


def hoja_gantt_oc(wb, ocs, calendario, talleres_orden, talleres):
    """Gantt por TALLER con una fila por OC y las unidades de cada día."""
    ws = wb.create_sheet("GANTT_OC")
    fijas = ["TALLER", "OC", "LÍNEAS", "MODELO", "COLOR", "CANTIDAD", "CAP/DÍA LÍNEA",
             "INICIO", "FIN", "FECHA FIN OC", "ESTADO"]
    c0 = encabezado_calendario(ws, fijas, calendario)
    fila = 4
    for taller in talleres_orden:
        del_taller = [o for o in ocs if o["TALLER"] == taller]
        totales = {}
        for o in del_taller:
            for k, u in o["POR_DIA"].items():
                totales[k] = totales.get(k, 0) + u
        fila_taller(ws, fila, c0, calendario, taller, del_taller, talleres.get(taller, (1, 1))[0], totales)
        ws.cell(row=fila, column=4, value=None)
        ws.cell(row=fila, column=6, value=sum(o["CANTIDAD"] for o in del_taller))
        fila += 1
        for n, o in enumerate(del_taller):
            valores = [taller, o["OC"], o["LINEAS"], o["MODELO"], o["COLOR"], o["CANTIDAD"], round(o["CAP_DIA"], 1),
                       o["INICIO PROG"], o["FIN PROG"], o["FECHA FIN OC"], o["ESTADO"]]
            for j, v in enumerate(valores, 1):
                c = ws.cell(row=fila, column=j, value=v)
                c.border = BORDE
                if j in (8, 9, 10):
                    c.number_format = FMT_FECHA
            ws.cell(row=fila, column=11).font = Font(bold=True, color="C00000" if o["ESTADO"] == "ATRASADA" else "00703C")
            for k in range(len(calendario)):
                c = ws.cell(row=fila, column=c0 + k)
                c.border = BORDE
                if k in o["POR_DIA"]:
                    tarde = o["FECHA FIN OC"] and calendario[k] > o["FECHA FIN OC"]
                    celda_barra(c, o["POR_DIA"][k], COLOR_ATRASO if tarde else COLORES_OC[n % len(COLORES_OC)])
            fila += 1
        fila += 1
    for j, a in enumerate([11, 7, 8, 24, 9, 9, 8, 11, 11, 11, 11], 1):
        ws.column_dimensions[get_column_letter(j)].width = a
    ws.cell(row=fila + 1, column=1, value="Número en cada celda = unidades de la OC ese día (sumando sus líneas). "
                                          "ROJO = después de la FECHA FIN de la OC.")


def guardar(ocs, calendario, feriados, origen_feriados, talleres, origen_talleres, fecha_inicio,
            advertencias, archivo_salida):
    talleres_orden = list(dict.fromkeys(o["TALLER"] for o in ocs))
    wb = Workbook()
    hoja_gantt_lineas(wb, ocs, calendario, talleres_orden, talleres)
    hoja_gantt_oc(wb, ocs, calendario, talleres_orden, talleres)

    ws = wb.create_sheet("PROGRAMA_OC")
    columnas = ["SECUENCIA", "TALLER", "OC", "LÍNEAS", "CLIENTE", "DESCRIPCION", "FAMILIA", "MODELO", "COLOR",
                "ESTADO OC", "TALLES", "CANTIDAD PLANIFICADA", "CAP/DÍA LÍNEA", "DÍAS HÁBILES",
                "INICIO PROGRAMADO", "FIN PROGRAMADO", "FECHA FIN OC", "ESTADO", "DÍAS HÁBILES DE ATRASO"]
    filas = []
    for t in talleres_orden:
        for i, o in enumerate([o for o in ocs if o["TALLER"] == t], 1):
            filas.append([i, t, o["OC"], o["LINEAS"], o["CLIENTE"], o["DESCRIPCION"], o["FAMILIA"], o["MODELO"],
                          o["COLOR"], o["ESTADO OC"], o["TALLES"], o["CANTIDAD"], round(o["CAP_DIA"], 1),
                          round(o["DIAS"], 2), o["INICIO PROG"], o["FIN PROG"], o["FECHA FIN OC"], o["ESTADO"],
                          o["ATRASO"]])
    tabla(ws, columnas, filas,
          anchos={"CLIENTE": 30, "DESCRIPCION": 30, "MODELO": 24, "INICIO PROGRAMADO": 13, "FIN PROGRAMADO": 13},
          formatos={"INICIO PROGRAMADO": FMT_FECHA, "FIN PROGRAMADO": FMT_FECHA, "FECHA FIN OC": FMT_FECHA})
    col_estado = columnas.index("ESTADO") + 1
    for r in range(2, len(filas) + 2):
        if ws.cell(row=r, column=col_estado).value == "ATRASADA":
            ws.cell(row=r, column=col_estado).fill = PatternFill("solid", fgColor="FFC7CE")

    ws = wb.create_sheet("DETALLE_DIARIO")
    filas = []
    for o in ocs:
        for t in o["TRAMOS"]:
            for k in sorted(t["POR_DIA"]):
                filas.append([calendario[k], o["TALLER"], nombre_linea(t["LINEA"]), o["OC"], o["MODELO"],
                              o["COLOR"], t["POR_DIA"][k]])
    filas.sort(key=lambda f: (talleres_orden.index(f[1]), f[0], f[2]))
    tabla(ws, ["FECHA", "TALLER", "LÍNEA", "OC", "MODELO", "COLOR", "UNIDADES"], filas,
          anchos={"FECHA": 12, "MODELO": 24}, formatos={"FECHA": FMT_FECHA})

    ws = wb.create_sheet("CARGA_LINEAS")
    filas = []
    for t in talleres_orden:
        del_t = [o for o in ocs if o["TALLER"] == t]
        n_lineas, max_oc = talleres.get(t, (1, 1))
        for i in range(n_lineas):
            tramos = [(o, tr) for o in del_t for tr in o["TRAMOS"] if tr["LINEA"] == i]
            if not tramos:
                filas.append([t, nombre_linea(i), 0, 0, 0, None, None])
                continue
            dias = sorted({k for _, tr in tramos for k in tr["POR_DIA"]})
            filas.append([t, nombre_linea(i), len(tramos), sum(tr["UNIDADES"] for _, tr in tramos),
                          round(sum(tr["FIN"] - tr["INICIO"] for _, tr in tramos), 2),
                          calendario[dias[0]], calendario[dias[-1]]])
    tabla(ws, ["TALLER", "LÍNEA", "N° OC", "UNIDADES", "DÍAS HÁBILES", "INICIO", "FIN"], filas,
          formatos={"INICIO": FMT_FECHA, "FIN": FMT_FECHA})

    ws = wb.create_sheet("RESUMEN_TALLER")
    filas = []
    for t in talleres_orden:
        del_t = [o for o in ocs if o["TALLER"] == t]
        n_lineas, max_oc = talleres.get(t, (1, 1))
        caps = sorted({round(o["CAP_DIA"], 1) for o in del_t})
        filas.append([t, n_lineas, max_oc, ", ".join(f"{c:g}" for c in caps),
                      ", ".join(f"{c * n_lineas:g}" for c in caps), len(del_t), sum(o["CANTIDAD"] for o in del_t),
                      min(o["INICIO PROG"] for o in del_t), max(o["FIN PROG"] for o in del_t),
                      sum(1 for o in del_t if o["ESTADO"] == "ATRASADA")])
    tabla(ws, ["TALLER", "LÍNEAS", "MÁX. LÍNEAS POR OC", "CAP/DÍA LÍNEA", "CAP/DÍA TALLER", "N° OC", "UNIDADES",
               "INICIO", "FIN", "OC ATRASADAS"], filas, formatos={"INICIO": FMT_FECHA, "FIN": FMT_FECHA})
    r = len(filas) + 3
    notas = [
        f"Fecha de inicio: {fecha_inicio:%d/%m/%Y}.",
        f"Capacidad de una línea = OPERADORES x {MINUTOS_POR_DIA} min / TIEMPO (hoja CAPACIDAD).",
        f"Líneas por taller: {origen_talleres}.",
        "Orden: el de la hoja LISTA OC; las OC se suman sin abrir por talle.",
        "Cada OC va a las líneas que se liberan primero (hasta el máximo por OC), repartida para que terminen a la vez.",
        "Una OC de menos de un día de una línea no se divide en dos líneas.",
        f"Feriados: {origen_feriados}.",
    ]
    for i, n in enumerate(notas):
        ws.cell(row=r + i, column=1, value=n)

    ws = wb.create_sheet("FERIADOS")
    hasta = calendario[-1] if calendario else fecha_inicio
    filas = [[f, d] for f, d in sorted(feriados.items()) if fecha_inicio <= f <= hasta + timedelta(days=31)]
    tabla(ws, ["FECHA", "DESCRIPCION"], filas, anchos={"FECHA": 12, "DESCRIPCION": 60}, formatos={"FECHA": FMT_FECHA})

    if advertencias:
        ws = wb.create_sheet("ADVERTENCIAS")
        tabla(ws, ["ADVERTENCIA"], [[a] for a in advertencias], anchos={"ADVERTENCIA": 100})

    os.makedirs(os.path.dirname(archivo_salida), exist_ok=True)
    wb.save(archivo_salida)


def main():
    archivo = sys.argv[1] if len(sys.argv) > 1 else ARCHIVO_ENTRADA
    fecha_inicio = date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else FECHA_INICIO

    advertencias = []
    cap = leer_capacidad(archivo)
    feriados, origen_feriados = leer_feriados(archivo)
    talleres, origen_talleres = leer_talleres(archivo)
    ocs = leer_ocs(archivo, cap, advertencias)
    if not ocs:
        print("No hay OC para programar.")
        return
    for t in dict.fromkeys(o["TALLER"] for o in ocs):
        if t not in talleres:
            advertencias.append(f"Taller {t}: no está en la tabla de talleres; se programa con 1 línea.")
    calendario = programar(ocs, fecha_inicio, feriados, talleres)
    guardar(ocs, calendario, feriados, origen_feriados, talleres, origen_talleres, fecha_inicio,
            advertencias, ARCHIVO_SALIDA)

    print(f"Programa generado: {ARCHIVO_SALIDA}")
    for t in dict.fromkeys(o["TALLER"] for o in ocs):
        del_t = [o for o in ocs if o["TALLER"] == t]
        print(f"  {t} ({talleres.get(t, (1, 1))[0]} línea/s): {len(del_t)} OC, {sum(o['CANTIDAD'] for o in del_t)} u, "
              f"{min(o['INICIO PROG'] for o in del_t):%d/%m/%Y} -> {max(o['FIN PROG'] for o in del_t):%d/%m/%Y}, "
              f"{sum(1 for o in del_t if o['ESTADO'] == 'ATRASADA')} atrasadas")
    for a in advertencias:
        print("  AVISO:", a)


if __name__ == "__main__":
    main()

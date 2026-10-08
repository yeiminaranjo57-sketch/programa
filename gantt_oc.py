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
  4. En cada TALLER pone las OC una detrás de otra, en el orden de la hoja
     LISTA OC, desde la fecha de inicio. Si una OC termina a mitad de día,
     la siguiente usa el resto de ese día.
  5. Solo cuenta días hábiles: sin sábados, domingos ni feriados.
     Los feriados se toman de la hoja "FERIADOS" del archivo de entrada si
     existe (columna FECHA); si no, se usa la lista de abajo (Argentina).
  6. Guarda "resultados/Programa_Gantt_OC.xlsx" con el Gantt por taller,
     el programa por OC, el detalle diario y el resumen por taller.

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

# Colores de las barras (uno por taller) y de alerta
COLORES_TALLER = ["4F81BD", "9BBB59", "F79646", "8064A2", "4BACC6", "C0504D", "2C4D75", "77933C"]
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


def programar(ocs, fecha_inicio, feriados):
    """Pone las OC en fila dentro de cada taller. El tiempo se mide en
    'días hábiles' con decimales: 2,5 = mitad del tercer día hábil."""
    puntero = {}
    for oc in ocs:
        inicio = puntero.get(oc["TALLER"], 0.0)
        duracion = oc["CANTIDAD"] / oc["CAP_DIA"]
        fin = inicio + duracion
        puntero[oc["TALLER"]] = fin
        oc["T_INICIO"], oc["T_FIN"], oc["DIAS"] = inicio, fin, duracion

        # Reparte las unidades por día; redondeo acumulado para que la suma
        # sea exactamente la cantidad de la OC.
        oc["POR_DIA"] = {}
        hecho = 0
        k = int(math.floor(inicio))
        while k < fin - 1e-9:
            hasta = min(fin, k + 1)
            acumulado = round((hasta - inicio) * oc["CAP_DIA"]) if hasta < fin else oc["CANTIDAD"]
            if acumulado - hecho > 0:
                oc["POR_DIA"][k] = acumulado - hecho
            hecho = acumulado
            k += 1

    total_dias = int(math.ceil(max(puntero.values()))) if puntero else 0
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


def hoja_gantt(wb, ocs, calendario, feriados, talleres):
    ws = wb.active
    ws.title = "GANTT_TALLER"
    fijas = ["TALLER", "OC", "MODELO", "COLOR", "CANTIDAD", "CAP/DÍA", "INICIO", "FIN", "FECHA FIN OC", "ESTADO"]
    c0 = len(fijas) + 1
    fino = Side(style="thin", color="BFBFBF")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
    dias_sem = ["LU", "MA", "MI", "JU", "VI"]

    # Fila 1: mes / Fila 2: día de la semana / Fila 3: fecha
    for j, nombre in enumerate(fijas, 1):
        ws.merge_cells(start_row=1, start_column=j, end_row=3, end_column=j)
        estilo_encabezado(ws.cell(row=1, column=j, value=nombre))
    inicio_mes = c0
    for k, d in enumerate(calendario):
        col = c0 + k
        estilo_encabezado(ws.cell(row=2, column=col, value=dias_sem[d.weekday()]))
        c = ws.cell(row=3, column=col, value=d)
        estilo_encabezado(c)
        c.number_format = "dd/mm"
        c.alignment = Alignment(horizontal="center", text_rotation=90)
        ultimo = k == len(calendario) - 1 or calendario[k + 1].month != d.month
        if ultimo:
            if col > inicio_mes:
                ws.merge_cells(start_row=1, start_column=inicio_mes, end_row=1, end_column=col)
            estilo_encabezado(ws.cell(row=1, column=inicio_mes, value=f"{meses[d.month - 1]} {d.year}"))
            inicio_mes = col + 1
    ws.row_dimensions[3].height = 42

    fila = 4
    for t_idx, taller in enumerate(talleres):
        color = COLORES_TALLER[t_idx % len(COLORES_TALLER)]
        del_taller = [o for o in ocs if o["TALLER"] == taller]

        # Fila resumen del taller: unidades totales por día
        ws.cell(row=fila, column=1, value=taller)
        ws.cell(row=fila, column=2, value=f"{len(del_taller)} OC")
        ws.cell(row=fila, column=5, value=sum(o["CANTIDAD"] for o in del_taller))
        ws.cell(row=fila, column=7, value=min(o["INICIO PROG"] for o in del_taller))
        ws.cell(row=fila, column=8, value=max(o["FIN PROG"] for o in del_taller))
        atrasadas = sum(1 for o in del_taller if o["ESTADO"] == "ATRASADA")
        ws.cell(row=fila, column=10, value=f"{atrasadas} atrasadas" if atrasadas else "OK")
        totales = {}
        for o in del_taller:
            for k, u in o["POR_DIA"].items():
                totales[k] = totales.get(k, 0) + u
        for k in range(len(calendario)):
            if k in totales:
                ws.cell(row=fila, column=c0 + k, value=totales[k])
        for col in range(1, c0 + len(calendario)):
            c = ws.cell(row=fila, column=col)
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor=COLOR_TALLER_FILA)
            c.border = borde
            if col in (7, 8):
                c.number_format = "dd/mm/yyyy"
            if col >= c0:
                c.alignment = Alignment(horizontal="center", text_rotation=90)
                c.font = Font(bold=True, size=8)
        fila += 1

        # Una fila por OC con su barra
        for o in del_taller:
            valores = [taller, o["OC"], o["MODELO"], o["COLOR"], o["CANTIDAD"], round(o["CAP_DIA"], 1),
                       o["INICIO PROG"], o["FIN PROG"], o["FECHA FIN OC"], o["ESTADO"]]
            for j, v in enumerate(valores, 1):
                c = ws.cell(row=fila, column=j, value=v)
                c.border = borde
                if j in (7, 8, 9):
                    c.number_format = "dd/mm/yyyy"
            estado = ws.cell(row=fila, column=10)
            estado.font = Font(bold=True, color="C00000" if o["ESTADO"] == "ATRASADA" else "00703C")
            for k in range(len(calendario)):
                c = ws.cell(row=fila, column=c0 + k)
                c.border = borde
                if k in o["POR_DIA"]:
                    tarde = o["FECHA FIN OC"] and calendario[k] > o["FECHA FIN OC"]
                    c.value = o["POR_DIA"][k]
                    c.fill = PatternFill("solid", fgColor=COLOR_ATRASO if tarde else color)
                    c.font = Font(color="FFFFFF", size=8, bold=True)
                    c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
            fila += 1
        fila += 1  # línea en blanco entre talleres

    anchos = [11, 7, 24, 9, 9, 8, 11, 11, 11, 11]
    for j, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = a
    for k in range(len(calendario)):
        ws.column_dimensions[get_column_letter(c0 + k)].width = 4.2
    ws.freeze_panes = ws.cell(row=4, column=c0)

    leyenda = fila + 1
    ws.cell(row=leyenda, column=1, value="Leyenda").font = Font(bold=True)
    ws.cell(row=leyenda + 1, column=1, value="Número en cada celda = unidades programadas ese día.")
    ws.cell(row=leyenda + 2, column=1, value="Barra en ROJO = días producidos después de la FECHA FIN de la OC.")
    ws.cell(row=leyenda + 3, column=1, value="Solo se muestran días hábiles: sin sábados, domingos ni feriados (ver hoja FERIADOS).")
    ws.cell(row=leyenda + 4, column=1, value="Fila azul clara = total del taller por día (debe coincidir con su capacidad diaria).")


def hoja_gantt_compacto(wb, ocs, calendario, talleres):
    """Una sola fila por taller: cada día muestra la(s) OC que se trabajan."""
    ws = wb.create_sheet("GANTT_COMPACTO", 1)
    fino = Side(style="thin", color="BFBFBF")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    estilo_encabezado(ws.cell(row=1, column=1, value="TALLER"))
    for k, d in enumerate(calendario):
        c = ws.cell(row=1, column=2 + k, value=d)
        estilo_encabezado(c)
        c.number_format = "dd/mm"
        c.alignment = Alignment(horizontal="center", text_rotation=90)
    ws.row_dimensions[1].height = 42
    for t_idx, taller in enumerate(talleres):
        fila_oc, fila_u = 2 + t_idx * 3, 3 + t_idx * 3
        ws.cell(row=fila_oc, column=1, value=taller).font = Font(bold=True)
        ws.cell(row=fila_u, column=1, value="unidades").font = Font(italic=True, size=8)
        base = COLORES_TALLER[t_idx % len(COLORES_TALLER)]
        del_taller = [o for o in ocs if o["TALLER"] == taller]
        por_dia = {}
        for n, o in enumerate(del_taller):
            for k, u in o["POR_DIA"].items():
                por_dia.setdefault(k, []).append((n, o, u))
        for k, lista in por_dia.items():
            n, o, _ = lista[-1]
            tarde = any(x[1]["FECHA FIN OC"] and calendario[k] > x[1]["FECHA FIN OC"] for x in lista)
            c = ws.cell(row=fila_oc, column=2 + k, value=" / ".join(str(x[1]["OC"]) for x in lista))
            # Se alterna tono claro/oscuro entre OC consecutivas para distinguirlas
            c.fill = PatternFill("solid", fgColor=COLOR_ATRASO if tarde else (base if n % 2 == 0 else "404040"))
            c.font = Font(color="FFFFFF", size=8, bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90, wrap_text=True)
            u = ws.cell(row=fila_u, column=2 + k, value=sum(x[2] for x in lista))
            u.font = Font(size=7)
            u.alignment = Alignment(horizontal="center", text_rotation=90)
        ws.row_dimensions[fila_oc].height = 70
        for col in range(1, 2 + len(calendario)):
            ws.cell(row=fila_oc, column=col).border = borde
    ws.column_dimensions["A"].width = 12
    for k in range(len(calendario)):
        ws.column_dimensions[get_column_letter(2 + k)].width = 4.2
    ws.freeze_panes = "B2"
    nota = 3 + len(talleres) * 3
    ws.cell(row=nota, column=1, value="Cada celda muestra la OC del día (dos OC si una termina y otra empieza ese día). "
                                      "Tonos alternados = OC distintas; ROJO = después de la FECHA FIN de la OC.")


def guardar(ocs, calendario, feriados, origen_feriados, fecha_inicio, advertencias, archivo_salida):
    talleres = list(dict.fromkeys(o["TALLER"] for o in ocs))
    wb = Workbook()
    hoja_gantt(wb, ocs, calendario, feriados, talleres)
    hoja_gantt_compacto(wb, ocs, calendario, talleres)

    fmt_f = "dd/mm/yyyy"
    ws = wb.create_sheet("PROGRAMA_OC")
    columnas = ["SECUENCIA", "TALLER", "OC", "CLIENTE", "DESCRIPCION", "FAMILIA", "MODELO", "COLOR", "ESTADO OC",
                "TALLES", "CANTIDAD PLANIFICADA", "CAPACIDAD/DÍA", "DÍAS HÁBILES", "INICIO PROGRAMADO",
                "FIN PROGRAMADO", "FECHA FIN OC", "ESTADO", "DÍAS HÁBILES DE ATRASO"]
    filas = []
    for t in talleres:
        for i, o in enumerate([o for o in ocs if o["TALLER"] == t], 1):
            filas.append([i, t, o["OC"], o["CLIENTE"], o["DESCRIPCION"], o["FAMILIA"], o["MODELO"], o["COLOR"],
                          o["ESTADO OC"], o["TALLES"], o["CANTIDAD"], round(o["CAP_DIA"], 1), round(o["DIAS"], 2),
                          o["INICIO PROG"], o["FIN PROG"], o["FECHA FIN OC"], o["ESTADO"], o["ATRASO"]])
    tabla(ws, columnas, filas,
          anchos={"CLIENTE": 30, "DESCRIPCION": 30, "MODELO": 24, "INICIO PROGRAMADO": 13, "FIN PROGRAMADO": 13},
          formatos={"INICIO PROGRAMADO": fmt_f, "FIN PROGRAMADO": fmt_f, "FECHA FIN OC": fmt_f})
    rojo = PatternFill("solid", fgColor="FFC7CE")
    for r in range(2, len(filas) + 2):
        if ws.cell(row=r, column=17).value == "ATRASADA":
            ws.cell(row=r, column=17).fill = rojo

    ws = wb.create_sheet("DETALLE_DIARIO")
    filas = []
    for o in ocs:
        for k in sorted(o["POR_DIA"]):
            filas.append([calendario[k], o["TALLER"], o["OC"], o["MODELO"], o["COLOR"], o["POR_DIA"][k]])
    filas.sort(key=lambda f: (talleres.index(f[1]), f[0]))
    tabla(ws, ["FECHA", "TALLER", "OC", "MODELO", "COLOR", "UNIDADES"], filas,
          anchos={"FECHA": 12, "MODELO": 24}, formatos={"FECHA": fmt_f})

    ws = wb.create_sheet("RESUMEN_TALLER")
    filas = []
    for t in talleres:
        del_t = [o for o in ocs if o["TALLER"] == t]
        filas.append([t, len(del_t), sum(o["CANTIDAD"] for o in del_t),
                      ", ".join(sorted({f"{o['CAP_DIA']:.1f}" for o in del_t})),
                      round(sum(o["DIAS"] for o in del_t), 2),
                      min(o["INICIO PROG"] for o in del_t), max(o["FIN PROG"] for o in del_t),
                      sum(1 for o in del_t if o["ESTADO"] == "ATRASADA")])
    tabla(ws, ["TALLER", "N° OC", "UNIDADES", "CAPACIDAD/DÍA", "DÍAS HÁBILES", "INICIO", "FIN", "OC ATRASADAS"],
          filas, formatos={"INICIO": fmt_f, "FIN": fmt_f})
    r = len(filas) + 3
    notas = [
        f"Fecha de inicio: {fecha_inicio:%d/%m/%Y}.",
        f"Capacidad diaria = OPERADORES x {MINUTOS_POR_DIA} min / TIEMPO (hoja CAPACIDAD).",
        "Orden dentro de cada taller: el mismo de la hoja LISTA OC; las OC se suman sin abrir por talle.",
        "Si una OC termina a mitad de día, la siguiente OC del taller empieza ese mismo día.",
        f"Feriados: {origen_feriados}.",
    ]
    for i, n in enumerate(notas):
        ws.cell(row=r + i, column=1, value=n)

    ws = wb.create_sheet("FERIADOS")
    hasta = calendario[-1] if calendario else fecha_inicio
    filas = [[f, d] for f, d in sorted(feriados.items()) if fecha_inicio <= f <= hasta + timedelta(days=31)]
    tabla(ws, ["FECHA", "DESCRIPCION"], filas, anchos={"FECHA": 12, "DESCRIPCION": 60}, formatos={"FECHA": fmt_f})

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
    feriados, origen = leer_feriados(archivo)
    ocs = leer_ocs(archivo, cap, advertencias)
    if not ocs:
        print("No hay OC para programar.")
        return
    calendario = programar(ocs, fecha_inicio, feriados)
    guardar(ocs, calendario, feriados, origen, fecha_inicio, advertencias, ARCHIVO_SALIDA)

    print(f"Programa generado: {ARCHIVO_SALIDA}")
    for t in dict.fromkeys(o["TALLER"] for o in ocs):
        del_t = [o for o in ocs if o["TALLER"] == t]
        print(f"  {t}: {len(del_t)} OC, {sum(o['CANTIDAD'] for o in del_t)} u, "
              f"{min(o['INICIO PROG'] for o in del_t):%d/%m/%Y} -> {max(o['FIN PROG'] for o in del_t):%d/%m/%Y}, "
              f"{sum(1 for o in del_t if o['ESTADO'] == 'ATRASADA')} atrasadas")
    for a in advertencias:
        print("  AVISO:", a)


if __name__ == "__main__":
    main()

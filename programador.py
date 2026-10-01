"""
PROGRAMADOR DE PRODUCCIÓN - Confección con procesos de decoración
=================================================================

Qué hace este archivo (en palabras sencillas):
  1. Lee "configuracion.xlsx": rutas de cada referencia, recursos
     (máquinas, módulos, personas) y de dónde sacar cada dato.
  2. Lee los archivos que usted descarga de SAP (Query Manager) y otros
     Excel, que deja en la carpeta "entradas":
        - Órdenes de corte    (qué se produce y cuánto)
        - Órdenes de venta    (cliente y fecha de entrega)
        - Explosión de materiales (si hay material para empezar)
  3. Une todo, ordena las órdenes (prioridad y fecha de entrega) y pasa
     cada una por sus procesos: Diseño -> Corte -> Sublimado / Bordado /
     Transfer / Pad print -> Confección -> Despacho.
  4. Guarda el resultado en "programa_produccion.xlsx" con semáforo,
     programa por recurso, carga, materiales faltantes y Gantt.

Cómo se usa:
  python programador.py
"""

import glob
import os
import sys
import unicodedata
from datetime import datetime, timedelta, time

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CONFIG = os.path.join(CARPETA, "configuracion.xlsx")
CARPETA_ENTRADAS = os.path.join(CARPETA, "entradas")
ARCHIVO_SALIDA = os.path.join(CARPETA, "programa_produccion.xlsx")

# Campos que el programa necesita de cada fuente (los marcados con * son obligatorios)
CAMPOS = {
    "Ordenes_corte": ["Orden_corte*", "Orden_venta", "Referencia*", "Cantidad*",
                      "Prioridad", "Fecha_entrega", "Descripcion"],
    "Ordenes_venta": ["Orden_venta*", "Cliente", "Fecha_entrega*"],
    "Materiales": ["Orden_corte*", "Material*", "Cantidad_requerida*",
                   "Cantidad_disponible*", "Fecha_llegada"],
}
FUENTES_OPCIONALES = {"Materiales"}


# ---------------------------------------------------------------------------
# 1. LECTURA DE ARCHIVOS (SAP y Excel varios)
# ---------------------------------------------------------------------------

def normalizar(texto):
    """'Fecha de Entrega ' -> 'fecha de entrega' (sin tildes ni mayúsculas)."""
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return " ".join(texto.lower().replace("_", " ").split())


def es_si(valor):
    return normalizar(valor) in ("si", "s", "x", "1", "true", "yes")


def buscar_archivo(patron):
    """
    Busca en la carpeta 'entradas'. Acepta comodines: 'ordenes_corte*' toma
    el archivo más reciente que empiece así (útil si SAP le pone la fecha al nombre).
    """
    candidatos = [f for f in glob.glob(os.path.join(CARPETA_ENTRADAS, patron))
                  if not os.path.basename(f).startswith("~$")]
    if not candidatos:
        candidatos = [f for f in glob.glob(os.path.join(CARPETA_ENTRADAS, patron + ".*"))
                      if not os.path.basename(f).startswith("~$")]
    return max(candidatos, key=os.path.getmtime) if candidatos else None


def leer_tabla(ruta, hoja, columnas_esperadas):
    """
    Lee .xlsx, .xls, .csv o .txt (tabulado, como exporta SAP).
    Si el archivo trae títulos arriba, busca sola la fila de encabezados.
    """
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".csv", ".txt"):
        with open(ruta, "rb") as f:
            inicio = f.read(4)
        codificacion = "utf-16" if inicio[:2] in (b"\xff\xfe", b"\xfe\xff") else \
            "utf-8-sig" if inicio[:3] == b"\xef\xbb\xbf" else "latin-1"
        crudo = pd.read_csv(ruta, sep="\t" if ext == ".txt" else None, header=None, dtype=str,
                            engine="python", encoding=codificacion)
    else:
        crudo = pd.read_excel(ruta, sheet_name=hoja if pd.notna(hoja) and hoja else 0, header=None)

    esperadas = {normalizar(c) for c in columnas_esperadas}
    fila_enc = 0
    for i in range(min(15, len(crudo))):
        valores = {normalizar(v) for v in crudo.iloc[i] if pd.notna(v)}
        if len(valores & esperadas) >= max(1, len(esperadas) // 2):
            fila_enc = i
            break
    tabla = crudo.iloc[fila_enc + 1:].copy()
    tabla.columns = [str(c).strip() for c in crudo.iloc[fila_enc]]
    return tabla.dropna(how="all")


def cargar_fuente(nombre, fuentes, advertencias):
    """Lee una fuente (ej. 'Ordenes_corte') y renombra sus columnas a las del programa."""
    reglas = fuentes[fuentes["Dato"] == nombre]
    if reglas.empty or reglas["Archivo"].isna().all():
        if nombre in FUENTES_OPCIONALES:
            return None
        sys.exit(f"ERROR: en configuracion.xlsx (hoja Fuentes) falta indicar el archivo de '{nombre}'.")

    patron = reglas["Archivo"].dropna().iloc[0]
    hoja = reglas["Hoja"].dropna().iloc[0] if reglas["Hoja"].notna().any() else None
    ruta = buscar_archivo(str(patron).strip())
    if ruta is None:
        if nombre in FUENTES_OPCIONALES:
            advertencias.append(("-", f"No se encontró el archivo de {nombre} ('{patron}'). "
                                      f"Se programó sin revisar materiales."))
            return None
        sys.exit(f"ERROR: no encuentro '{patron}' en la carpeta 'entradas'.")

    # Campo del programa -> nombre de la columna en su archivo
    mapa = {r["Campo_programa"]: r["Columna_en_su_archivo"] for _, r in reglas.iterrows()
            if pd.notna(r["Campo_programa"]) and pd.notna(r["Columna_en_su_archivo"])}
    tabla = leer_tabla(ruta, hoja, mapa.values())
    por_nombre = {normalizar(c): c for c in tabla.columns}

    datos = pd.DataFrame(index=tabla.index)
    for campo in CAMPOS[nombre]:
        obligatorio, campo = campo.endswith("*"), campo.rstrip("*")
        columna = por_nombre.get(normalizar(mapa.get(campo, campo)))
        if columna is not None:
            datos[campo] = tabla[columna]
        elif obligatorio:
            sys.exit(f"ERROR: en '{os.path.basename(ruta)}' no encuentro la columna "
                     f"'{mapa.get(campo, campo)}' (campo {campo}). Revise la hoja Fuentes.")
    print(f"  {nombre:<14} <- {os.path.basename(ruta)} ({len(datos)} filas)")
    return datos


def a_fecha(serie):
    return pd.to_datetime(serie, dayfirst=True, errors="coerce")


def a_numero(serie):
    """Convierte '1.200' o '1,200.00' o '1200,5' en número."""
    def convertir(v):
        if pd.isna(v) or isinstance(v, (int, float)):
            return v
        s = str(v).strip().replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        return pd.to_numeric(s, errors="coerce")
    return serie.map(convertir).astype(float)


def texto(serie):
    """Códigos como texto, sin el '.0' que a veces agrega Excel."""
    return serie.map(lambda v: "" if pd.isna(v) else
                     str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip())


def preparar_ordenes(oc, ov, advertencias):
    oc["Orden_corte"] = texto(oc["Orden_corte"])
    oc["Referencia"] = texto(oc["Referencia"])
    oc["Cantidad"] = a_numero(oc["Cantidad"])
    oc["Prioridad"] = a_numero(oc["Prioridad"]).fillna(3) if "Prioridad" in oc else 3.0
    oc["Fecha_entrega"] = a_fecha(oc["Fecha_entrega"]) if "Fecha_entrega" in oc else pd.NaT

    # Traer cliente y fecha de entrega desde la orden de venta
    if "Orden_venta" in oc:
        oc["Orden_venta"] = texto(oc["Orden_venta"])
        ov = ov.copy()
        ov["Orden_venta"] = texto(ov["Orden_venta"])
        ov["Fecha_entrega"] = a_fecha(ov["Fecha_entrega"])
        ov = ov.drop_duplicates("Orden_venta").rename(columns={"Fecha_entrega": "Fecha_entrega_ov"})
        oc = oc.merge(ov, on="Orden_venta", how="left")
        oc["Fecha_entrega"] = oc["Fecha_entrega"].fillna(oc["Fecha_entrega_ov"])
    if "Cliente" not in oc:
        oc["Cliente"] = ""
    if "Descripcion" not in oc:
        oc["Descripcion"] = ""

    sin_fecha = oc["Fecha_entrega"].isna()
    for _, o in oc[sin_fecha].iterrows():
        advertencias.append((o["Orden_corte"], "Sin fecha de entrega (no se encontró su orden de venta). "
                                               "Se programó al final."))
    oc["Fecha_entrega"] = oc["Fecha_entrega"].fillna(pd.Timestamp.max.normalize())
    oc = oc[oc["Cantidad"] > 0]
    return oc.drop_duplicates("Orden_corte")


def revisar_materiales(mat, inicio, advertencias):
    """
    Devuelve, por orden de corte, desde cuándo hay material completo.
    None = falta material y no hay fecha de llegada (la orden queda bloqueada).
    """
    if mat is None:
        return {}, pd.DataFrame()
    mat = mat.copy()
    mat["Orden_corte"] = texto(mat["Orden_corte"])
    mat["Cantidad_requerida"] = a_numero(mat["Cantidad_requerida"]).fillna(0)
    mat["Cantidad_disponible"] = a_numero(mat["Cantidad_disponible"]).fillna(0)
    mat["Fecha_llegada"] = a_fecha(mat["Fecha_llegada"]) if "Fecha_llegada" in mat else pd.NaT

    faltantes = mat[mat["Cantidad_disponible"] < mat["Cantidad_requerida"]].copy()
    faltantes["Faltante"] = faltantes["Cantidad_requerida"] - faltantes["Cantidad_disponible"]

    listo = {}
    for orden, filas in faltantes.groupby("Orden_corte"):
        if filas["Fecha_llegada"].isna().any():
            listo[orden] = None
        else:
            listo[orden] = max(inicio, filas["Fecha_llegada"].max().to_pydatetime())
    columnas = ["Orden_corte", "Material", "Cantidad_requerida", "Cantidad_disponible",
                "Faltante", "Fecha_llegada"]
    return listo, faltantes[columnas]


# ---------------------------------------------------------------------------
# 2. CALENDARIO DE CADA RECURSO (máquina, módulo o persona)
# ---------------------------------------------------------------------------

class Recurso:
    def __init__(self, fila, inicio_programa, hora_turno):
        self.nombre = str(fila["Recurso"]).strip()
        self.proceso = normalizar(fila["Proceso"])
        self.horas_dia = min(float(fila.get("Horas_por_dia", 8) or 8), 24)
        self.sabado = es_si(fila.get("Trabaja_sabado", "No"))
        self.domingo = es_si(fila.get("Trabaja_domingo", "No"))
        self.hora_turno = hora_turno
        disponible = fila.get("Disponible_desde")
        self.disponible_desde = max(inicio_programa, pd.to_datetime(disponible).to_pydatetime()) \
            if pd.notna(disponible) else inicio_programa
        self.ocupado = []  # lista de (inicio, fin) de los trabajos ya asignados

    def trabaja_el_dia(self, fecha):
        dia = fecha.weekday()  # 0 = lunes ... 5 = sábado, 6 = domingo
        if dia == 5:
            return self.sabado
        if dia == 6:
            return self.domingo
        return True

    def sumar_horas_trabajo(self, desde, horas):
        """
        Devuelve (inicio_real, fin) de un trabajo de 'horas' que empieza
        no antes de 'desde', saltando noches y días no laborables.
        """
        actual, inicio_real, restante = desde, None, horas
        # Revisamos desde el día anterior por si hay un turno que cruza la medianoche
        dia = datetime.combine(desde.date(), time(0)) - timedelta(days=1)
        while True:
            if self.trabaja_el_dia(dia):
                ini = datetime.combine(dia.date(), time(self.hora_turno))
                fin = ini + timedelta(hours=self.horas_dia)
                if fin > actual:
                    actual = max(actual, ini)
                    if inicio_real is None:
                        inicio_real = actual
                    disponible = (fin - actual).total_seconds() / 3600
                    if restante <= disponible:
                        return inicio_real, actual + timedelta(hours=restante)
                    restante -= disponible
                    actual = fin
            dia += timedelta(days=1)

    def primer_hueco(self, listo_desde, horas):
        """
        Busca el primer espacio libre donde cabe el trabajo completo.
        Así, si el recurso queda esperando por un pedido sin material,
        otra orden puede aprovechar ese tiempo muerto.
        """
        desde = max(listo_desde, self.disponible_desde)
        intentos = [desde] + sorted(f for _, f in self.ocupado if f > desde)
        for intento in intentos:
            ini, fin = self.sumar_horas_trabajo(intento, horas)
            if all(fin <= o_ini or ini >= o_fin for o_ini, o_fin in self.ocupado):
                return ini, fin

    def reservar(self, ini, fin):
        self.ocupado.append((ini, fin))


# ---------------------------------------------------------------------------
# 3. PROGRAMACIÓN (el "cerebro")
# ---------------------------------------------------------------------------

def numero(valor, defecto=0.0):
    return float(valor) if pd.notna(valor) and str(valor).strip() != "" else defecto


def programar(ordenes, rutas, recursos_df, inicio, hora_turno, material_listo, advertencias):
    recursos = [Recurso(f, inicio, hora_turno) for _, f in recursos_df.iterrows()]
    programa = []
    bloqueadas = set()

    # Regla: primero prioridad (1 = más urgente), luego la fecha de entrega más cercana
    for _, orden in ordenes.sort_values(["Prioridad", "Fecha_entrega"]).iterrows():
        oc = orden["Orden_corte"]
        pasos = rutas[rutas["Referencia"] == orden["Referencia"]].sort_values("Paso")
        if pasos.empty:
            advertencias.append((oc, f"La referencia {orden['Referencia']} no tiene ruta en "
                                     f"configuracion.xlsx. No se programó."))
            continue

        listo_desde = inicio
        for _, paso in pasos.iterrows():
            # La regla del material aplica desde el paso marcado (normalmente Corte)
            if es_si(paso.get("Requiere_material", "No")) and oc in material_listo:
                if material_listo[oc] is None:
                    advertencias.append((oc, f"Falta material sin fecha de llegada. Quedó "
                                             f"programada solo hasta antes de {paso['Proceso']}."))
                    bloqueadas.add(oc)
                    break
                listo_desde = max(listo_desde, material_listo[oc])

            proceso = normalizar(paso["Proceso"])
            candidatos = [r for r in recursos if r.proceso == proceso]
            permitidos = paso.get("Recursos_permitidos")
            if pd.notna(permitidos) and str(permitidos).strip():
                nombres = {n.strip() for n in str(permitidos).split(";")}
                candidatos = [r for r in candidatos if r.nombre in nombres]
            if not candidatos:
                advertencias.append((oc, f"No hay recursos para el proceso '{paso['Proceso']}'. "
                                         f"La orden quedó incompleta."))
                break

            # Tiempo = alistamiento + horas fijas + unidades / velocidad
            velocidad = numero(paso.get("Unidades_por_hora"))
            horas = numero(paso.get("Horas_alistamiento")) + numero(paso.get("Horas_fijas"))
            if velocidad > 0:
                horas += orden["Cantidad"] / velocidad
            horas = max(horas, 0.01)

            # Probamos todos los recursos posibles y nos quedamos con el que termina primero
            mejor = None
            for r in candidatos:
                ini, fin = r.primer_hueco(listo_desde, horas)
                if mejor is None or fin < mejor[2]:
                    mejor = (r, ini, fin)
            r, ini, fin = mejor
            r.reservar(ini, fin)

            programa.append({
                "Orden_corte": oc,
                "Orden_venta": orden.get("Orden_venta", ""),
                "Cliente": orden["Cliente"],
                "Referencia": orden["Referencia"],
                "Cantidad": orden["Cantidad"],
                "Paso": paso["Paso"],
                "Proceso": paso["Proceso"],
                "Recurso": r.nombre,
                "Inicio": ini,
                "Fin": fin,
                "Horas": round(horas, 1),
                "Fecha_entrega": orden["Fecha_entrega"],
            })
            listo_desde = fin + timedelta(hours=numero(paso.get("Horas_espera_despues")))

    return pd.DataFrame(programa), bloqueadas


def resumen_por_orden(programa, ordenes, bloqueadas):
    filas = []
    for _, o in ordenes.sort_values(["Prioridad", "Fecha_entrega"]).iterrows():
        oc = o["Orden_corte"]
        ops = programa[programa["Orden_corte"] == oc] if not programa.empty else programa
        entrega = o["Fecha_entrega"]
        fila = {"Orden_corte": oc, "Orden_venta": o.get("Orden_venta", ""), "Cliente": o["Cliente"],
                "Referencia": o["Referencia"], "Cantidad": o["Cantidad"], "Prioridad": o["Prioridad"],
                "Fecha_entrega": entrega.date() if entrega.year < 2200 else None,
                "Inicio": None, "Fin_estimado": None, "Dias_holgura": None}
        if oc in bloqueadas:
            fila["Estado"] = "FALTA MATERIAL"
        elif ops.empty:
            fila["Estado"] = "SIN PROGRAMAR"
        else:
            fin = ops["Fin"].max()
            holgura = (entrega + timedelta(days=1) - fin).total_seconds() / 86400
            fila.update(Inicio=ops["Inicio"].min(), Fin_estimado=fin,
                        Dias_holgura=round(holgura, 1) if entrega.year < 2200 else None)
            fila["Estado"] = "ATRASADO" if holgura < 0 else "EN RIESGO" if holgura < 2 else "A TIEMPO"
        filas.append(fila)
    return pd.DataFrame(filas)


def orden_de_procesos(programa):
    """Procesos en el orden en que aparecen en las rutas (Diseño, Corte, ...)."""
    return {p: i for i, p in enumerate(programa.groupby("Proceso")["Paso"].min().sort_values().index)}


def carga_por_recurso(programa):
    if programa.empty:
        return pd.DataFrame()
    carga = programa.groupby(["Proceso", "Recurso"]).agg(
        Ordenes=("Orden_corte", "count"), Unidades=("Cantidad", "sum"),
        Horas_programadas=("Horas", "sum"), Ocupado_hasta=("Fin", "max")).reset_index()
    carga["_orden"] = carga["Proceso"].map(orden_de_procesos(programa))
    return carga.sort_values(["_orden", "Recurso"]).drop(columns="_orden")


# ---------------------------------------------------------------------------
# 4. GUARDAR EN EXCEL CON FORMATO
# ---------------------------------------------------------------------------

COLORES_ESTADO = {"ATRASADO": "F8696B", "EN RIESGO": "FFEB84", "A TIEMPO": "63BE7B",
                  "FALTA MATERIAL": "F4B183", "SIN PROGRAMAR": "BFBFBF"}
COLORES_PROCESO = ["4F81BD", "F79646", "9BBB59", "8064A2", "4BACC6", "C0504D", "2C4D75", "D9A441"]


def ajustar_hoja(ws):
    for celda in ws[1]:
        celda.fill = PatternFill("solid", fgColor="1F3864")
        celda.font = Font(color="FFFFFF", bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        largo = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(largo + 2, 10), 60)
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            if isinstance(c.value, datetime):
                c.number_format = "dd/mm/yyyy hh:mm"
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def dibujar_gantt(wb, programa, inicio):
    """Hoja con una fila por recurso y una columna por día, coloreada por proceso."""
    ws = wb.create_sheet("Gantt")
    if programa.empty:
        return
    dia0 = inicio.date()
    dias = [dia0 + timedelta(days=i) for i in range((programa["Fin"].max().date() - dia0).days + 1)]
    orden_proc = orden_de_procesos(programa)
    color = {p: COLORES_PROCESO[i % len(COLORES_PROCESO)] for p, i in orden_proc.items()}

    ws.cell(1, 1, "Recurso")
    ws.column_dimensions["A"].width = 26
    for j, d in enumerate(dias, start=2):
        ws.cell(1, j, d.strftime("%a %d/%m"))
        ws.column_dimensions[get_column_letter(j)].width = 9

    grupos = sorted(programa.groupby(["Proceso", "Recurso"]), key=lambda g: (orden_proc[g[0][0]], g[0][1]))
    for i, ((proceso, recurso), ops) in enumerate(grupos, start=2):
        ws.cell(i, 1, f"{recurso} ({proceso})")
        for _, op in ops.iterrows():
            for j, d in enumerate(dias, start=2):
                if op["Inicio"].date() <= d <= op["Fin"].date():
                    c = ws.cell(i, j)
                    c.fill = PatternFill("solid", fgColor=color[proceso])
                    c.value = f"{c.value}, {op['Orden_corte']}" if c.value else str(op["Orden_corte"])
                    c.font = Font(color="FFFFFF", size=8)

    for celda in ws[1]:
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "B2"


def guardar(resumen, programa, carga, faltantes, advertencias, inicio):
    hojas = {
        "Resumen_ordenes": resumen,
        "Programa_por_recurso": programa.assign(_o=programa["Proceso"].map(orden_de_procesos(programa)))
        .sort_values(["_o", "Recurso", "Inicio"]).drop(columns=["_o", "Paso"]) if not programa.empty else programa,
        "Programa_por_orden": programa.sort_values(["Orden_corte", "Paso"]) if not programa.empty else programa,
        "Carga_recursos": carga,
        "Materiales_faltantes": faltantes,
        "Advertencias": pd.DataFrame(advertencias, columns=["Orden_corte", "Advertencia"]),
    }
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as xl:
        for nombre, df in hojas.items():
            df.to_excel(xl, sheet_name=nombre, index=False)

    wb = load_workbook(ARCHIVO_SALIDA)
    for ws in wb.worksheets:
        ajustar_hoja(ws)
    ws = wb["Resumen_ordenes"]
    col = [c.value for c in ws[1]].index("Estado") + 1
    for fila in range(2, ws.max_row + 1):
        c = ws.cell(fila, col)
        if c.value in COLORES_ESTADO:
            c.fill = PatternFill("solid", fgColor=COLORES_ESTADO[c.value])
            c.font = Font(bold=True)
    dibujar_gantt(wb, programa, inicio)
    wb.save(ARCHIVO_SALIDA)


# ---------------------------------------------------------------------------
# 5. PROGRAMA PRINCIPAL
# ---------------------------------------------------------------------------

def main():
    print("Leyendo configuracion.xlsx ...")
    config = pd.read_excel(ARCHIVO_CONFIG, sheet_name=None)
    rutas = config["Rutas"].dropna(subset=["Referencia", "Proceso"])
    rutas["Referencia"] = texto(rutas["Referencia"])
    recursos = config["Recursos"].dropna(subset=["Recurso", "Proceso"])
    fuentes = config["Fuentes"]
    param = dict(zip(config["Parametros"]["Parametro"], config["Parametros"]["Valor"]))
    inicio = pd.to_datetime(param.get("Fecha_inicio_programa", datetime.today()))
    if pd.isna(inicio):
        inicio = pd.Timestamp.today()
    hora_turno = int(numero(param.get("Hora_inicio_turno"), 6))
    inicio = datetime.combine(inicio.date(), time(hora_turno))

    print("Leyendo archivos de la carpeta 'entradas' ...")
    advertencias = []
    oc = cargar_fuente("Ordenes_corte", fuentes, advertencias)
    ov = cargar_fuente("Ordenes_venta", fuentes, advertencias)
    mat = cargar_fuente("Materiales", fuentes, advertencias)

    ordenes = preparar_ordenes(oc, ov, advertencias)
    material_listo, faltantes = revisar_materiales(mat, inicio, advertencias)
    programa, bloqueadas = programar(ordenes, rutas, recursos, inicio, hora_turno,
                                     material_listo, advertencias)
    resumen = resumen_por_orden(programa, ordenes, bloqueadas)
    carga = carga_por_recurso(programa)

    try:
        guardar(resumen, programa, carga, faltantes, advertencias, inicio)
    except PermissionError:
        sys.exit("ERROR: cierre programa_produccion.xlsx en Excel y vuelva a intentarlo.")

    print(f"\nResultado ({len(ordenes)} órdenes de corte):")
    for estado, n in resumen["Estado"].value_counts().items():
        print(f"  {estado:<15} {n}")
    if advertencias:
        print(f"  ¡Atención! {len(advertencias)} advertencias (ver hoja 'Advertencias').")
    print("\nListo. Abra el archivo: programa_produccion.xlsx")


if __name__ == "__main__":
    main()

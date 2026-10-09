# Programador automático de producción — Confección y decoración

Guía pensada para alguien que **no es programador**. Se lee en 10 minutos.

---

## 1. La idea en una frase

Hoy usted descarga varias queries de SAP B1, las cruza con otros Excel y arma el
programa a mano. Esta herramienta **hace el cruce y arma el programa por usted**,
siguiendo reglas claras, en segundos.

> **¿Esto es Inteligencia Artificial?** No exactamente, y está bien que no lo sea.
> Es **automatización con reglas**: el computador sigue instrucciones claras
> ("primero lo urgente", "no cortar si no hay tela"). Es confiable y usted puede
> revisar cada decisión. La IA (como Claude) sirve para **ayudarle a construir y
> ajustar** la herramienta conversando en español.

## 2. Una analogía: la receta de cocina

| En la cocina | En su programa de producción |
|---|---|
| Ingredientes | Sus archivos: órdenes de venta, órdenes de corte, explosión de materiales |
| Despensa fija | `configuracion.xlsx`: rutas, recursos, horarios |
| Receta | Las reglas de programación |
| Cocinero | `programador.py` (usted no necesita abrirlo) |
| Plato servido | `programa_produccion.xlsx` |

## 3. Cómo fluye la información

```
 SAP B1 (Query Manager)        Otros Excel
   ├─ Órdenes de venta    ─┐
   ├─ Órdenes de corte    ─┼──►  carpeta "entradas"  ──►  programador  ──►  programa_produccion.xlsx
   └─ Explosión materiales ┘            ▲
                                        │
                          configuracion.xlsx (rutas, recursos, fuentes)
```

El programa **une** los archivos usando los números de documento:

- **Orden de corte → Orden de venta**: de ahí saca el *cliente* y la *fecha de entrega*.
- **Orden de corte → Explosión de materiales**: revisa si hay tela e insumos.
  Si falta material y se sabe cuándo llega, no programa el corte antes de esa fecha.
  Si no hay fecha de llegada, marca la orden como **FALTA MATERIAL**.

## 4. Las reglas que sigue

1. **Orden de atención**: primero *Prioridad* (1 = urgente), luego la *fecha de entrega* más cercana.
2. **Ruta de cada referencia**: por ejemplo
   - Camiseta sublimada: Diseño → Sublimado → Corte → Confección → Despacho
   - Polo bordado: Diseño → Corte → Bordado → Confección → Despacho
   - Buzo con transfer: Diseño → Corte → Confección → Transfer → Despacho
   - Camiseta con pad print: Diseño → Corte → Pad print → Confección → Despacho

   Un proceso no empieza hasta que termina el anterior (más horas de espera si las hay,
   por ejemplo secado).
3. **Tiempo de cada proceso** = alistamiento + horas fijas + (prendas ÷ prendas por hora).
   - Diseño y Despacho suelen ser *horas fijas por orden* (ej. 3 h).
   - Corte, Bordado, Confección, etc. dependen de la *cantidad*.
4. **Recurso**: en cada proceso usa la máquina, módulo o persona que **termine primero**.
   Si un recurso tiene un hueco libre (por ejemplo, porque otra orden espera tela),
   lo aprovecha.
5. **Horarios**: respeta las horas por día, sábados, domingos y fechas de mantenimiento.
6. **Semáforo**: 🟢 A TIEMPO · 🟡 EN RIESGO (menos de 2 días de margen) · 🔴 ATRASADO ·
   🟠 FALTA MATERIAL.

## 5. Los archivos

### `configuracion.xlsx` (se llena una vez y se ajusta de vez en cuando)

| Hoja | Qué contiene |
|---|---|
| **Fuentes** | En qué archivo está cada dato y **cómo se llama la columna en SU archivo**. |
| **Rutas** | Procesos de cada referencia, en orden (Paso 10, 20, 30…), con sus tiempos. |
| **Recursos** | Diseñadores, mesas de corte, sublimadora, bordadoras, planchas de transfer, pad print, módulos de confección, despacho, con horarios. |
| **Parametros** | Fecha y hora de inicio del programa. |

**La hoja Fuentes es la clave para trabajar con SAP + otros Excel.** Ejemplo:

| Dato | Campo_programa | Columna_en_su_archivo | Archivo |
|---|---|---|---|
| Ordenes_venta | Orden_venta | DocNum | ordenes_venta* |
| Ordenes_venta | Cliente | CardName | ordenes_venta* |
| Ordenes_venta | Fecha_entrega | DocDueDate | ordenes_venta* |

Se lee así: *"en el archivo que empieza por `ordenes_venta`, la columna `DocNum` es el número de la orden de venta"*.
Si mañana cambian una query en SAP, **solo se cambia esta tabla**; no hay que tocar nada más.

Detalles que facilitan la vida:
- El `*` en *Archivo* toma **el archivo más reciente** que empiece así, por ejemplo `ordenes_corte_2026-10-05.xlsx`.
  No tiene que renombrar las descargas.
- No importan las mayúsculas ni las tildes en los nombres de columnas.
- Si SAP pone un título encima de la tabla, el programa encuentra solo la fila de encabezados.
- Lee `.xlsx`, `.xls`, `.csv` y `.txt` (texto tabulado).
- Acepta fechas `dd/mm/aaaa` y números como `1.200,50`.

### Carpeta `entradas` (se actualiza cada vez que programa)

Ahí deja lo que descarga del Query Manager y los otros Excel. Hoy trae **archivos de
ejemplo** con nombres de columnas parecidos a los de SAP; reemplácelos por los suyos.

| Dato | Obligatorio | Campos que necesita |
|---|---|---|
| Órdenes de corte | Sí | N° orden de corte, N° orden de venta, referencia, cantidad (opcional: prioridad, descripción) |
| Órdenes de venta | Sí | N° orden de venta, cliente, fecha de entrega |
| Explosión de materiales | No | N° orden de corte, material, requerido, disponible (opcional: fecha de llegada) |

## 6. Instalación (una sola vez, en Windows)

1. Instale **Python** desde <https://www.python.org/downloads/>.
   En la primera pantalla **marque "Add Python to PATH"**.
2. Descargue esta carpeta a su computador.
3. Doble clic en **`instalar.bat`**.

## 7. Uso diario

1. Ejecute sus queries en SAP, exporte a Excel y guarde los archivos en la carpeta **`entradas`**.
2. Si cambió la fecha de inicio, actualícela en `configuracion.xlsx` → *Parametros*.
3. Doble clic en **`ejecutar.bat`**.
4. Abra **`programa_produccion.xlsx`**:

| Hoja | Para qué sirve |
|---|---|
| **Resumen_ordenes** | Semáforo por orden de corte: cuándo termina y si cumple |
| **Programa_por_recurso** | Lista de trabajo de cada máquina, módulo o persona (para entregar a planta) |
| **Programa_por_orden** | Recorrido de cada orden por todos sus procesos |
| **Carga_recursos** | Horas asignadas a cada recurso; muestra el **cuello de botella** |
| **Materiales_faltantes** | Qué falta, para qué orden y cuándo llega (útil para compras) |
| **Advertencias** | Referencias sin ruta, órdenes sin fecha, material sin fecha de llegada |
| **Gantt** | Calendario visual: una fila por recurso, una columna por día |

## 8. Cómo leer el resultado

- **Orden en rojo**: mire en *Programa_por_orden* en qué proceso se demora. Opciones:
  subirle la prioridad a 1, agregar horas (subir *Horas_por_dia*), habilitar el sábado
  o permitir más recursos.
- **Cuello de botella**: en *Carga_recursos*, el recurso con más horas o con la fecha
  *Ocupado_hasta* más lejana. En el ejemplo es **Confección** (los módulos).
- **Simular escenarios**: cambie un dato, vuelva a ejecutar y compare.
  Por ejemplo: *"¿y si el módulo 3 trabaja el sábado?"*.

## 9. Límites actuales (y cómo crecer)

Esta versión **todavía no**:

- Procesa en **paralelo** dos decoraciones de la misma orden (ej. bordado y sublimado
  sobre piezas distintas al mismo tiempo); hoy van una después de la otra.
- Divide una orden grande entre varios módulos de confección a la vez.
- Agrupa órdenes con el mismo diseño o los mismos colores para ahorrar alistamientos.
- Reparte el material escaso entre órdenes que compiten por la misma tela.

Camino recomendado:

1. **Fase 1 (esta):** conectar sus archivos reales (llenar la hoja *Fuentes*) y usarla
   en paralelo al programa manual durante 2 a 3 semanas para comparar.
2. **Fase 2:** ajustar reglas según lo que se observe (paralelos, división de órdenes, lotes).
3. **Fase 3:** que las queries corran solas (conexión directa a SAP), sin descargar a mano.

## 10. Glosario rápido

- **Query**: consulta guardada en SAP que trae una tabla de datos.
- **Ruta**: secuencia de procesos que sigue una referencia.
- **Recurso**: quien hace el trabajo: una máquina, un módulo de confección o una persona.
- **Cuello de botella**: el proceso más cargado, que limita a toda la planta.
- **Holgura**: días de margen entre el fin estimado y la fecha de entrega.
- **Gantt**: calendario de barras; cada barra es un trabajo en un recurso.

---

## 11. Gantt de OC por taller y línea (`gantt_oc.py`)

Programa las OC de la hoja **LISTA OC** (sumando todos los talles de cada OC) en su taller,
usando la hoja **CAPACIDAD** del mismo archivo.

- **Entrada:** `entradas/PRUEBA_GANTT.xlsx` (hojas `LISTA OC` y `CAPACIDAD`).
- **Capacidad de UNA línea** = OPERADORES × 495 min ÷ TIEMPO, buscada por MODELO + COLOR.
- **Líneas por taller:** CASEROS tiene 7 líneas y cada OC puede ir como máximo en 2;
  OLIDEN tiene 1 línea. Para cambiarlo, agregue al archivo de entrada una hoja `TALLERES`
  con las columnas `TALLER`, `LINEAS` y `MAX_LINEAS_POR_OC`.
- **Orden:** el de la hoja LISTA OC. Cada OC va a las líneas que se liberan primero
  (hasta el máximo), con la cantidad repartida para que terminen a la vez. Una OC de menos
  de un día de una línea no se divide. Si una OC termina a mitad de día, la siguiente
  empieza ese mismo día en esa línea.
- **Calendario:** desde el 09/10/2026, solo días hábiles: sin sábados, domingos ni
  feriados nacionales de Argentina. Para usar otros feriados, agregue al archivo de
  entrada una hoja `FERIADOS` con la columna `FECHA` (y opcional `DESCRIPCION`).
- **Uso:** doble clic en `ejecutar_gantt.bat`, o `python gantt_oc.py [archivo] [AAAA-MM-DD]`.
- **Salida:** `resultados/Programa_Gantt_OC.xlsx`

| Hoja | Para qué sirve |
|---|---|
| **GANTT_LINEAS** | Gantt por taller: una fila por línea con la OC de cada día y sus unidades. Rojo = después de la FECHA FIN |
| **GANTT_OC** | Gantt por taller con una fila por OC, sus líneas y las unidades de cada día |
| **PROGRAMA_OC** | Líneas, inicio y fin programados de cada OC, estado y días de atraso |
| **DETALLE_DIARIO** | Fecha / taller / línea / OC / unidades (para tablas dinámicas) |
| **CARGA_LINEAS** | Unidades, días de carga y fecha de fin de cada línea |
| **RESUMEN_TALLER** | Líneas, capacidad del taller, unidades y fecha de fin por taller |
| **FERIADOS** | Feriados que se descontaron |

# Programador automático de producción — Tejido de punto

Guía pensada para alguien que **no es programador**. Se lee en 10 minutos.

---

## 1. La idea en una frase

Hoy usted descarga datos de SAP B1 y arma el programa a mano en Excel.
Este proyecto hace que **el computador arme el programa por usted**, siguiendo
las mismas reglas que usted ya aplica, en segundos.

> **¿Esto es Inteligencia Artificial?** No exactamente, y está bien que no lo sea.
> Es **automatización con reglas**: el computador sigue instrucciones claras
> ("primero lo urgente", "usar la máquina que termine antes"). Es más confiable y
> fácil de revisar que una IA. La IA (como Claude) sirve para **ayudarle a construir
> y ajustar** esta herramienta conversando en español, sin que usted programe.

## 2. Una analogía: la receta de cocina

| En la cocina | En su programa de producción |
|---|---|
| Ingredientes | Los datos: pedidos (de SAP), rutas, máquinas |
| Receta (pasos) | Las reglas de programación |
| Cocinero | El archivo `programador.py` |
| Plato servido | El Excel `programa_produccion.xlsx` |

Usted solo cambia los **ingredientes** (pega los pedidos nuevos) y el cocinero
prepara el plato. Si quiere otro sabor, se ajusta la **receta** (las reglas).

## 3. Qué hace, paso a paso

```
 SAP B1  ──►  datos_entrada.xlsx  ──►  programador.py  ──►  programa_produccion.xlsx
(descarga)     (usted pega aquí)        (doble clic)          (resultado con colores)
```

Las reglas que sigue hoy (se pueden cambiar):

1. **Ordena los pedidos**: primero por *Prioridad* (1 = urgente), luego por la
   *fecha de entrega* más cercana.
2. **Recorre los procesos** de cada artículo en orden
   (ej.: Tejido → Tintorería → Acabado → Revisión). Un proceso no empieza hasta
   que termina el anterior (más las horas de espera/reposo si las hay).
3. **Escoge máquina**: entre las máquinas de ese proceso (y solo las permitidas
   para ese artículo, ej. por diámetro o galga), usa la que **termine primero**.
4. **Calcula tiempos**: `horas = alistamiento + kilos ÷ kg por hora`, respetando
   el horario de cada máquina (horas por día, sábados, domingos).
5. **Avisa**: marca cada pedido como **A TIEMPO** (verde), **EN RIESGO** (amarillo,
   menos de 2 días de margen) o **ATRASADO** (rojo).

## 4. El archivo de entrada `datos_entrada.xlsx`

Tiene 4 hojas (más una de instrucciones). Viene con **datos de ejemplo** para que
lo pruebe; luego los reemplaza por los reales.

| Hoja | ¿Cada cuánto se actualiza? | Qué contiene |
|---|---|---|
| **Pedidos** | Cada vez que programa | Lo que descarga de SAP: Pedido, Articulo, Descripcion, Cliente, Cantidad_kg, Fecha_entrega, Prioridad |
| **Rutas** | Solo con artículos nuevos | Por artículo: Paso, Proceso, Kg_por_hora, Horas_alistamiento, Maquinas_permitidas (separadas por `;`), Horas_espera_despues |
| **Maquinas** | Cuando cambia la planta | Maquina, Proceso, Horas_por_dia, Trabaja_sabado, Trabaja_domingo, Disponible_desde (para mantenimientos) |
| **Parametros** | Cada vez que programa | Fecha_inicio_programa, Hora_inicio_turno |

**Importante:** los nombres de las columnas deben quedar **exactamente** como están.
El orden de las columnas no importa, y puede tener columnas extra.

## 5. De dónde sacar los datos en SAP B1

Lo más útil son las **órdenes de fabricación abiertas**. Si su área de sistemas
puede crear una consulta (Herramientas → Consultas → Generador/Asistente de consultas),
esta trae exactamente las columnas que necesita la hoja *Pedidos*:

```sql
SELECT  T0.DocNum                       AS "Pedido",
        T0.ItemCode                     AS "Articulo",
        T1.ItemName                     AS "Descripcion",
        T0.CardCode                     AS "Cliente",
        T0.PlannedQty - T0.CmpltQty     AS "Cantidad_kg",
        T0.DueDate                      AS "Fecha_entrega",
        3                               AS "Prioridad"
FROM    OWOR T0
JOIN    OITM T1 ON T1.ItemCode = T0.ItemCode
WHERE   T0.Status IN ('P','R')          -- P = planificada, R = liberada
ORDER BY T0.DueDate
```

Si programan desde **pedidos de venta** en vez de órdenes de fabricación, las tablas
son `ORDR` (encabezado) y `RDR1` (líneas, con `OpenQty` como cantidad pendiente).
Las tablas `OITT`/`ITT1` (listas de materiales) y, si las usan, las *rutas de producción*
de SAP B1 pueden alimentar la hoja *Rutas*.

Al exportar a Excel desde SAP, solo copie y pegue en la hoja *Pedidos*.

## 6. Instalación (una sola vez, en Windows)

1. Instale **Python** desde <https://www.python.org/downloads/>.
   En la primera pantalla **marque la casilla "Add Python to PATH"**.
2. Descargue esta carpeta a su computador.
3. Doble clic en **`instalar.bat`** (instala dos complementos: `pandas` y `openpyxl`).

## 7. Uso diario

1. Abra `datos_entrada.xlsx`, pegue los pedidos nuevos y actualice la fecha de inicio.
   Guarde y **cierre** el archivo.
2. Doble clic en **`ejecutar.bat`**.
3. Abra **`programa_produccion.xlsx`**:

| Hoja | Para qué sirve |
|---|---|
| **Resumen_pedidos** | Semáforo por pedido: cuándo termina y si cumple la fecha |
| **Programa_por_maquina** | La lista de trabajo de cada máquina, en orden (para entregar a planta) |
| **Carga_maquinas** | Horas y kilos asignados a cada máquina → muestra el **cuello de botella** |
| **Advertencias** | Artículos sin ruta, procesos sin máquina, etc. |
| **Gantt** | Calendario visual: una fila por máquina, una columna por día |

## 8. Cómo leer el resultado (con los datos de ejemplo)

- Si un pedido sale en **rojo**, mire en *Programa_por_maquina* en qué proceso se
  demora. Opciones: subirle la prioridad a 1, adelantar la fecha de inicio, agregar
  horas extra (subir *Horas_por_dia*) o permitir más máquinas.
- En *Carga_maquinas* la máquina con la fecha *Ocupada_hasta* más lejana es su
  **cuello de botella**. En el ejemplo es **Revisión** (solo trabaja 8 h/día):
  aunque tejido y tintorería terminen rápido, los pedidos se represan ahí.
- Cambie un dato, vuelva a ejecutar y compare: así puede **simular escenarios**
  ("¿qué pasa si la revisión trabaja 16 horas?") en segundos.

## 9. Límites actuales (y cómo crecer)

Esta es una **primera versión**. Hoy **no** tiene en cuenta:

- Agrupar pedidos del mismo color en tintorería (lotes por color / capacidad del jet).
- Disponibilidad de hilo o materia prima.
- Dividir un pedido grande entre varias máquinas al mismo tiempo.
- Tiempos de cambio que dependan del orden (ej. de color oscuro a claro).

Cada uno se puede agregar después. Camino recomendado:

1. **Fase 1 (esta):** probarla con datos reales en paralelo al programa manual
   durante 2–3 semanas y comparar.
2. **Fase 2:** ajustar reglas (lotes por color, materia prima, turnos reales).
3. **Fase 3:** que la descarga de SAP sea automática (consulta guardada o conexión
   directa a la base de datos), para no copiar y pegar.

## 10. Glosario rápido

- **Python**: el "idioma" en que está escrita la receta. Usted no necesita aprenderlo.
- **Script** (`programador.py`): el archivo con la receta.
- **Ruta**: la secuencia de procesos que sigue un artículo.
- **Cuello de botella**: el proceso más lento, que limita a toda la planta.
- **Holgura**: días de margen entre la fecha en que termina un pedido y su entrega.
- **Gantt**: gráfico de barras en el tiempo; cada barra es un trabajo en una máquina.

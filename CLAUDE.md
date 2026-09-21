# Routine: Registro semanal Meta Ads → Excel OneDrive

## Objetivo
Al ejecutarse en cualquier día, registrar en la hoja `REGISTRO_SEMANAL` del Excel de OneDrive la inversión, leads y conversaciones de Meta Ads para la semana actual (lunes → hoy) y, si falta, también la semana anterior completa (lunes → domingo). Una fila por campaña por semana. Campañas de leads (`OUTCOME_LEADS`) registran leads; campañas de WhatsApp (`OUTCOME_ENGAGEMENT`) registran conversaciones iniciadas.

Como métrica norte, también cuenta los **leads calientes** y **leads tibios** de la semana desde los archivos Excel de seguimiento de cada campaña en la carpeta MKT CODE de OneDrive. Un lead caliente/tibio es cualquier registro en la hoja `CALIENTE`/`TIBIO` de esos archivos cuya fecha caiga dentro del periodo procesado. El conteo se escribe en las columnas J (`Leads_Calientes`) y K (`Leads_Tibios`) de `REGISTRO_SEMANAL`. Cuando un mismo `Codigo_Producto` tiene varias filas de campaña en la misma semana, el conteo se escribe **solo en la primera fila** del grupo; las demás variantes reciben `0` — así la suma de las columnas J y K nunca duplica leads.

## Reglas

1. **Autonomía total**: no preguntes, no pidas confirmación, no esperes input. Si algo falla, detente con mensaje claro y exit ≠ 0.
2. **Una fila por campaña por semana** — nunca agrupar campañas aunque compartan código.
3. **Idempotencia por campaña**: si ya existe una fila con el mismo `ID_Campana` (nombre completo) + `Semana_ISO`, actualiza `Inversion_USD`, `Leads` y `Conversaciones`. No insertes duplicado.
4. **Semana anterior**: solo verifica y rellena la semana inmediatamente anterior (W-1). No retrocedas más.
5. **Filtrar campañas sin código**: si el nombre de la campaña no contiene el patrón `PE.EI.XX-XX.X` o `CE.EI.XX-XX.X`, omítela silenciosamente.
6. **CPL siempre como fórmula**, nunca valor literal.
7. Usa **Meta MCP** para leer datos de campañas. Usa **Composio Excel MCP** para leer y escribir el Excel.
8. **Escritura secuencial**: nunca escribas múltiples rangos en paralelo al mismo Excel — las escrituras simultáneas causan pérdida de datos. Escribe un bloque, espera confirmación, luego el siguiente.
9. **Solo campañas 2026+**: el filtro `amount_spent > 0` junto con el `time_range` semanal ya limita a campañas con actividad reciente. No se procesan campañas anteriores a 2026.
10. **Verificación anti-corrimiento (obligatoria)**: antes de actualizar `G{n}:I{n}` de una fila existente, leer primero `E{n}` y confirmar que coincide exactamente con el `ID_Campana` de la campaña que se va a escribir. Si no coincide, NO escribir: volver a leer el used range completo, recalcular el número de fila y reintentar. Después de terminar todas las escrituras de una semana, releer el bloque completo de esa semana y validar: (a) cada fila tiene A-F no vacíos, (b) el par E↔G:I corresponde a la campaña correcta según los datos de Meta, (c) no existe ninguna fila con G:I con valores pero A-F vacíos (fila fantasma). Si se detecta una fila fantasma, limpiarla con `EXCEL_CLEAR_RANGE` y corregir la fila que debió recibir esos datos.
11. **Recalcular fila tras cada inserción**: `primera_fila_vacia` no se debe mantener solo como contador en memoria — después de cada inserción, confirmar con la respuesta de `EXCEL_UPDATE_RANGE` que la fila escrita es la esperada (la respuesta incluye `address`). Si el Excel fue modificado por otro proceso durante la ejecución (used range distinto al inicial), releer el used range antes de seguir escribiendo.
12. **Respetar filas de otros canales**: al leer `filas_existentes`, ignorar para todo propósito de actualización las filas donde columna F (`Canal`) sea distinto de `"Meta"` (ej. `"WhatsApp Masivo"`, `"LinkedIn"`). Esas filas son de entrada manual — no tocarlas, no sobreescribirlas, no usarlas para calcular `ultimo_id` ni `primera_fila_vacia`.
13. **Consulta a Meta y campo `results`**: la consulta general con `fields: ["id","name","objective","amount_spent","lead","results"]` devuelve conversaciones en `results.values[0].value` también para campañas de mensajería (verificado 2026-09-04). Usar esa consulta como fuente primaria. Solo si `results` viene `{}` o `"Not available"` para una campaña de mensajería con spend > 0, ejecutar la consulta de respaldo con `breakdowns: ["publisher_platform"]` y sumar por plataforma (pasos 2a-bis / 3a-bis). No ejecutar la segunda consulta por defecto.
14. **Tres objetivos, no dos**: las campañas de mensajería pueden venir con `objective` = `OUTCOME_ENGAGEMENT` **o** `OUTCOME_SALES` (ej. `PE.EI.06-26.4 WHATSAPP` es OUTCOME_SALES). NUNCA clasificar por objetivo solo. La regla correcta es por el **indicador de `results`**:
    - `results.indicator` contiene `messaging_conversation_started` → es campaña de conversaciones: `leads = 0`, `conversaciones = results.values[0].value`
    - `results.indicator` contiene `leadgen` → es campaña de leads: `leads = lead`, `conversaciones = 0`
    Si se clasifica por `objective`, las campañas OUTCOME_SALES quedan sin registrar y su gasto se pierde.
15. **Campañas renombradas a mitad de semana**: Meta permite renombrar una campaña sin cambiar su `id`. Si la rutina corre el miércoles y otra vez el viernes, una campaña renombrada genera DOS filas (nombre viejo + nombre nuevo) para la misma semana, duplicando el gasto. Antes de insertar una fila nueva, verificar si ya existe otra fila de la misma semana con el mismo `Codigo_Producto` cuyo `ID_Campana` no aparezca en la respuesta actual de Meta: si es así, esa fila es un nombre obsoleto → sobreescribirla con el nombre nuevo en lugar de insertar. Toda fila de la semana procesada cuyo `ID_Campana` no exista en la respuesta de Meta es candidata a borrado (ver PASO 5-bis).
16. **Valores parciales vs finales**: correr la rutina a mitad de semana escribe datos parciales. Cada corrida posterior DEBE actualizar `G:I` de todas las filas de la semana en curso con los valores acumulados actuales, no solo insertar las campañas nuevas. Comparar fila por fila contra Meta y reescribir cualquier diferencia.

## Asignación MCP vs script

| Operación | Método |
|---|---|
| Consultar campañas, spend y leads de Meta Ads | Meta MCP (`ads_get_ad_entities`) |
| Leer hoja `REGISTRO_SEMANAL` | Composio `EXCEL_GET_WORKSHEET_USED_RANGE` |
| Actualizar celdas de fila existente | Composio `EXCEL_UPDATE_RANGE` |
| Insertar nuevas filas | Composio `EXCEL_UPDATE_RANGE` |
| Buscar archivo de campaña en MKT CODE | Composio `ONE_DRIVE_LIST_FOLDER_CHILDREN` (NO usar `EXCEL_SEARCH_FILES` — no alcanza drives ajenos) |
| Leer hojas `CALIENTE` y `TIBIO` de archivo de campaña | Composio `EXCEL_GET_WORKSHEET_USED_RANGE` |
| Leer "Cuadro de ventas diarias" (josecardenas, SOLO LECTURA) | MCP Microsoft `read_resource` — **NUNCA escribir** |
| Escribir ventas en TABLERO MKT-VENTAS | Composio `EXCEL_UPDATE_RANGE` |

---

## Pasos de ejecución

### PASO 0 — Calcular fechas

A partir de la fecha de hoy, calcular:

```
hoy              = fecha actual de ejecución  (ej. 2026-07-08, miércoles)

semana_actual_iso    = número ISO de semana de hoy  (ej. "2026-W28")
semana_actual_lunes  = lunes de esta semana          (ej. 2026-07-06)
semana_actual_hasta  = hoy                           (ej. 2026-07-08)

semana_anterior_iso    = semana_actual_iso - 1       (ej. "2026-W27")
semana_anterior_lunes  = lunes de la semana anterior (ej. 2026-06-29)
semana_anterior_domingo = semana_actual_lunes - 1 día (ej. 2026-07-05)
```

Para el serial Excel de una fecha (columna B): días transcurridos desde 1899-12-30.
- Ej: 2026-07-06 → 46209

---

### PASO 1 — Leer estado actual del Excel

Usa Composio `EXCEL_GET_WORKSHEET_USED_RANGE`:
- `item_id`: env `ONEDRIVE_ITEM_ID`
- `drive_id`: env `ONEDRIVE_DRIVE_ID`
- `worksheet_id`: `REGISTRO_SEMANAL`

De la respuesta (`.values`), construir en memoria:
- `ultimo_id` = máximo valor de la columna A (ID_Registro)
- `filas_existentes` = lista de todas las filas con sus valores
- `semanas_registradas` = set de valores únicos en columna C (Semana_ISO)
- `primera_fila_vacia` = número de fila Excel donde insertar la próxima fila nueva (= cantidad de filas de datos + 2, por el encabezado en fila 1)

---

### PASO 2 — Verificar y rellenar/actualizar semana anterior

Hay dos casos:

**Caso A — W-1 NO está en `semanas_registradas`** → insertar filas nuevas (la semana anterior nunca fue registrada).

**Caso B — W-1 SÍ está en `semanas_registradas` Y hoy es lunes** → actualizar las filas existentes de W-1 con los datos finales completos (lunes→domingo). La semana anterior ya cerró y los datos son definitivos.

Si W-1 ya está registrada y hoy no es lunes → omitir PASO 2 por completo.

En cualquiera de los dos casos activos:

**2a. Consultar Meta Ads para la semana anterior — campañas de LEADS**

Usa Meta MCP `ads_get_ad_entities`:
```
ad_account_id: env META_AD_ACCOUNT_ID
level: "campaign"
fields: ["id", "name", "objective", "amount_spent", "lead", "results"]
filtering: [{"field": "campaign.amount_spent", "operator": "GREATER_THAN", "value": ["0"]}]
time_range: {"since": "semana_anterior_lunes", "until": "semana_anterior_domingo"}
```

De esta consulta extraer spend y leads para campañas `OUTCOME_LEADS`. Para campañas `OUTCOME_ENGAGEMENT`, solo tomar el spend (NO usar `results` ni `lead` — ambos vienen vacíos/null).

**2a-bis. Consultar Meta Ads para la semana anterior — conversaciones WhatsApp (OBLIGATORIO)**

El campo `results` devuelve `{}` vacío para campañas `OUTCOME_ENGAGEMENT` en la consulta sin breakdown. Esta segunda consulta es **obligatoria** para obtener las conversaciones:

```
ad_account_id: env META_AD_ACCOUNT_ID
level: "campaign"
fields: ["id", "name", "objective", "amount_spent", "results"]
filtering: [
  {"field": "campaign.name", "operator": "CONTAIN", "value": ["Whatsapp"]},
  {"field": "campaign.amount_spent", "operator": "GREATER_THAN", "value": ["0"]}
]
time_range: {"since": "semana_anterior_lunes", "until": "semana_anterior_domingo"}
breakdowns: ["publisher_platform"]
```

Esto devuelve varias filas por campaña (una por plataforma: facebook, instagram, whatsapp). El campo `results.value` viene como `"N (Messaging conversations started)"`. Para cada campaña, sumar el número N de todas las plataformas:
```
conversaciones = sum(int(row['results']['value'].split()[0]) for row in rows_de_esta_campana)
```

**2b. Filtrar y procesar campañas de la semana anterior**

Por cada campaña en los resultados:
- Si `amount_spent` = 0 y `lead` = 0 y conversaciones = 0 → omitir
- Extraer `codigo` con regex `(PE\.EI\.\d+-\d+\.\d+|CE\.EI\.\d+-\d+\.\d+)` del campo `name`
- Si no hay coincidencia con PE.EI o CE.EI → omitir silenciosamente
- `id_campana` = `name` completo tal como viene de Meta
- Si `objective` == `OUTCOME_ENGAGEMENT` → campaña WhatsApp: `leads = 0`, `conversaciones` = total sumado del paso 2a-bis
- Si `objective` == `OUTCOME_LEADS` → campaña de leads: `leads` = valor de `lead`, `conversaciones = 0`

**2c. Escribir filas de la semana anterior**

Para cada campaña válida, insertar una fila nueva (no hay duplicados posibles porque la semana anterior no estaba en el Excel).

Ver estructura de columnas en la sección **Estructura de filas** más abajo.

Actualizar `ultimo_id` y `primera_fila_vacia` a medida que se insertan filas.

---

### PASO 3 — Procesar semana actual

**3a. Consultar Meta Ads para la semana actual — campañas de LEADS**

Usa Meta MCP `ads_get_ad_entities`:
```
ad_account_id: env META_AD_ACCOUNT_ID
level: "campaign"
fields: ["id", "name", "objective", "amount_spent", "lead", "results"]
filtering: [{"field": "campaign.amount_spent", "operator": "GREATER_THAN", "value": ["0"]}]
time_range: {"since": "semana_actual_lunes", "until": "semana_actual_hasta"}
```

**3a-bis. Consultar Meta Ads para la semana actual — conversaciones WhatsApp (OBLIGATORIO)**

Misma lógica que 2a-bis. Esta consulta es **obligatoria** — sin ella las campañas WhatsApp quedan con 0 conversaciones:

```
ad_account_id: env META_AD_ACCOUNT_ID
level: "campaign"
fields: ["id", "name", "objective", "amount_spent", "results"]
filtering: [
  {"field": "campaign.name", "operator": "CONTAIN", "value": ["Whatsapp"]},
  {"field": "campaign.amount_spent", "operator": "GREATER_THAN", "value": ["0"]}
]
time_range: {"since": "semana_actual_lunes", "until": "semana_actual_hasta"}
breakdowns: ["publisher_platform"]
```

Para cada campaña, sumar `int(results['value'].split()[0])` de todas las plataformas = total conversaciones.

**3b. Filtrar campañas**

Por cada campaña:
- Si `amount_spent` = 0 y `lead` = 0 y conversaciones = 0 → omitir
- Extraer `codigo` con regex del campo `name` → si no tiene PE.EI o CE.EI → omitir
- `id_campana` = `name` completo
- Si `objective` == `OUTCOME_ENGAGEMENT` → campaña WhatsApp: `leads = 0`, `conversaciones` = total sumado del paso 3a-bis
- Si `objective` == `OUTCOME_LEADS` → campaña de leads: `leads` = valor de `lead`, `conversaciones = 0`

**3c. Decidir: ¿insertar o actualizar?**

Para cada campaña válida, buscar en `filas_existentes` si existe una fila donde:
- Columna C (`Semana_ISO`) == `semana_actual_iso`
- Columna E (`ID_Campana`) == `campaign_name` (exacto)

**Si la fila EXISTE → ACTUALIZAR**
Primero verificar con `EXCEL_GET_RANGE` que `E{n}` == nombre de la campaña (regla 10). Luego usa Composio `EXCEL_UPDATE_RANGE` en las celdas G, H e I de esa fila:
- `address`: `G{n}:I{n}` (n = número de fila Excel de la coincidencia)
- `values`: [[spend, leads, conversaciones]]

**Si la fila NO EXISTE → INSERTAR**
Insertar nueva fila (ver estructura abajo) en `primera_fila_vacia`.
Incrementar `ultimo_id` y `primera_fila_vacia`.

---

### PASO 4 — Leads calientes y tibios por campaña (métrica norte)

Ejecutar **después** de que PASO 2 y PASO 3 hayan terminado de escribir todas las filas.

**4a. Construir mapa de códigos únicos a contar**

De las filas procesadas en la semana actual (y la anterior si aplicó backfill), extraer el set de `Codigo_Producto` únicos que fueron insertados o actualizados en esta ejecución.

**4b-0. De dónde sale el consolidado según el intake (CRÍTICO)**

**MKT CODE no contiene todos los consolidados.** Buscar solo ahí es la causa de que el Intake 2 apareciera con ceros durante semanas.

- **Intake 1** → carpeta MKT CODE en el drive de Daniel (`MKT_CODE_DRIVE_ID` / `MKT_CODE_FOLDER_ID`).
- **Intake 2** → la ruta de cada programa está declarada en **`routine-amg-semanal/programas.json`**, que es la fuente autoritativa. Leer ese archivo, no adivinar. Cada entrada trae un campo `consolidado` con prefijo de drive:
  - `daniel:/AREAS/MKT/MKT CODE/<archivo>.xlsx` → drive de Daniel, resolver el `id` listando la carpeta y **comparando el nombre completo del archivo**, no solo el prefijo del código.
  - `ti:<itemId>` → drive de TI (`b!oTTu_ShdmUi2TgpxUQHPzeseyr_zLFhBrnOvfZBsgzj3QOKX2iTWSpZPsgJ-Blho`), usar el itemId tal cual.

  De los 13 programas del Intake 2, **7 viven en el drive de TI y 6 en MKT CODE**.

**Nunca elegir por "el primer archivo cuyo nombre empieza con el código".** En MKT CODE conviven varios archivos por código (ej. hay 3 archivos `PE.EI.26`, y el correcto es `PE.EI.26-26.1 Renovable a la Red.xlsx`, no `PE.EI.26-26.1 Conexio a la red PROVISIONAL.xlsx`). El criterio es el nombre exacto que declara `programas.json`; en su defecto, el que coincida con el nombre de la campaña de Meta.

**El Intake 2 tiene 13 programas, no 11.** Faltaban en el catálogo de este archivo:
- `PE.EI.34-26.2` — BESS Colombia del Intake 2. Corre **en paralelo** con `PE.EI.34-26.1`, que es del Intake 1. La campaña de Meta `PE.EI.34-26.1 BESS COLOMBIA v4` es del Intake 1 y NO cuenta para el Intake 2. Comparar siempre el código **con edición completa**.
- `SM.EI.03-26.1` — Summit 2. **En Meta la campaña se llama `SM.IE.03-26.1`** (typo de origen, `IE` en vez de `EI`). El regex de extracción de código debe aceptar el prefijo `SM` además de `PE` y `CE`, y contemplar ese typo, o el gasto del Summit queda fuera del registro.

**4b. Para cada código único:**

Buscar el archivo de seguimiento en MKT CODE listando la carpeta con Composio `ONE_DRIVE_LIST_FOLDER_CHILDREN`. **El parámetro es `folder_item_id`, NO `item_id`** (con `item_id` la llamada falla o devuelve otra cosa):
```
folder_item_id: env MKT_CODE_FOLDER_ID    (= "01XEGP2KYCIGWDO6XTTVBZDOF2IIBDMQ5S")
drive_id: env MKT_CODE_DRIVE_ID
top: 200
select: ["id", "name", "lastModifiedDateTime"]
```

**Paginar siempre**: la carpeta tiene ~99 archivos y crece. La respuesta trae `next_page_token`; si viene con valor hay que repetir la llamada pasándolo como `page_token` y acumular, hasta que sea `null`. Sin paginar se pierden archivos silenciosamente.

NO usar `EXCEL_SEARCH_FILES` — solo busca en `/me/drive` (Natalie) y no alcanza la carpeta de Daniel.

Del listado, filtrar el archivo cuyo `name` empiece con el `codigo_producto` (ej. nombre empieza con `"PE.EI.34-26.1"`). Usar su `id` como `item_id` para las lecturas siguientes.

- Si no se encuentra ningún archivo → `leads_calientes = 0`, `leads_tibios = 0`, y **registrar el código en la lista de avisos** (ver 4e). No es un error, pero sí algo que un humano debe ver.
- Si hay varios resultados → usar el que tenga el código al inicio del nombre (coincidencia más específica).
- El listado se puede cachear en memoria para la ejecución completa — no es necesario listar la carpeta una vez por cada código.
- Guardar también `lastModifiedDateTime` de cada archivo: si un archivo no se modificó dentro de la semana procesada, sus conteos serán 0 legítimamente y conviene reportarlo como archivo estancado en lugar de como cero silencioso.

**4c. Leer hojas CALIENTE y TIBIO del archivo encontrado**

Usa Composio `EXCEL_GET_WORKSHEET_USED_RANGE` dos veces (secuencialmente):
```
item_id: item_id del archivo encontrado
drive_id: env MKT_CODE_DRIVE_ID
worksheet_id: "CALIENTE"   (primera llamada)
worksheet_id: "TIBIO"      (segunda llamada)
```

**Localizar la columna de fecha por NOMBRE DE ENCABEZADO, nunca por posición.** El orden de columnas cambia entre archivos — fijar la columna F es la causa de conteos en cero silenciosos. Igual que hace `routine-leads-frios/scripts/consolidar.py`:

```python
def norm(v):
    s = unicodedata.normalize("NFKD", str(v or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())

enc  = [norm(c) for c in filas[0]]
i_fecha = next((i for i, h in enumerate(enc) if h.startswith("fecha")), None)
i_tipo  = enc.index("tipo") if "tipo" in enc else None
if i_fecha is None:
    → aviso "archivo X: sin columna FECHA", conteo 0, continuar
```

**La temperatura la da LA HOJA, no la columna TIPO.** Es la regla fijada en `routine-amg-semanal/CLAUDE.md`, que es la autoridad para el Intake 2. Una fila en la pestaña `CALIENTE` es caliente aunque su columna TIPO diga otra cosa. No filtrar por TIPO.

**Exclusiones obligatorias** (sin ellas los conteos salen inflados):
1. **Leads internos**: descartar si el correo contiene `@inelinc.com` o el nombre contiene `PRUEBA`.
2. **Lead Inválido**: si alguna de las columnas `status final de atención` / `observaciones` contiene "invalido"/"inválido", la fila cuenta en el total pero **NO** en calientes ni tibios. Llevar el conteo aparte.

**Nombres de columna que hay que reconocer** (varían entre archivos):
- fecha: `fecha ingreso lead` o `fecha`
- nombre: `nombre completo` o `nombres completos`
- correo: `correo`

Una vez localizada la columna de fecha, su contenido sigue siendo inconsistente fila por fila:

- **Seriales numéricos** (ej. `46150.44`): Excel malparseó la fecha americana `mm/dd/yyyy` como `dd/mm/yyyy`, intercambiando día y mes. Ej: Meta envía `08/05/2026` (Aug 5) → Excel lee día=08 mes=05 → serial de **May 8** (46150). Esto pasa para TODOS los días 1-12 del mes.
- **Strings de texto** (ej. `"13/08/2026 10:33:29 AM"`): cuando el día del mes > 12, Excel no puede parsear `08/13/2026` como dd/mm → lo almacena como texto. Estos son correctos en formato `dd/mm/yyyy`.

**Parseo con tabla de seriales misread (método determinístico)**

Antes de procesar filas, precomputar la tabla de seriales "falsos" para el mes actual. La fórmula es: para cada día `DD` del mes (1-12), el serial misread = serial de `date(año, DD, mes_actual)` — porque Excel intercambió día↔mes.

Tabla para agosto 2026 (serial Excel = días desde 1899-12-30):

| Día real agosto | Meta envía | Excel lee como | Serial almacenado |
|---|---|---|---|
| Aug 1  | 08/01/2026 | 08/Ene → Jan 8  | 46030 |
| Aug 2  | 08/02/2026 | 08/Feb → Feb 8  | 46061 |
| Aug 3  | 08/03/2026 | 08/Mar → Mar 8  | 46089 |
| Aug 4  | 08/04/2026 | 08/Abr → Apr 8  | 46120 |
| Aug 5  | 08/05/2026 | 08/May → May 8  | 46150 |
| Aug 6  | 08/06/2026 | 08/Jun → Jun 8  | 46181 |
| Aug 7  | 08/07/2026 | 08/Jul → Jul 8  | 46211 |
| Aug 8  | 08/08/2026 | 08/Ago → Aug 8  | 46242 (¡este es correcto!) |
| Aug 9  | 08/09/2026 | 08/Sep → Sep 8  | 46273 |
| Aug 10 | 08/10/2026 | 08/Oct → Oct 8  | 46303 |
| Aug 11 | 08/11/2026 | 08/Nov → Nov 8  | 46334 |
| Aug 12 | 08/12/2026 | 08/Dic → Dec 8  | 46364 |

Para generalizar a cualquier mes M del año Y:
```
for dd in range(1, 13):
    misread_date = date(Y, dd, M)   # día↔mes intercambiados
    serial = (misread_date - date(1899, 12, 30)).days
    misread_serial_to_real_day[serial] = dd   # dd es el día real del mes M
```

**Algoritmo de parseo por fila:**

1. Para cada fila (saltando el encabezado), leer el valor de la columna F:
   - **Si es numérico** (int o float): truncar la parte decimal (hora). Buscar `int(val)` en la tabla `misread_serial_to_real_day`. Si hay match → la fecha real es `date(Y, M, día_real)`. Si no hay match → convertir serial normalmente (`date(1899,12,30) + timedelta(days=int(val))`) y usar esa fecha.
   - **Si es string**: extraer `A/B/YYYY` con regex `(\d{1,2})/(\d{1,2})/(\d{4})`. Interpretar como `dd/mm/yyyy` (A=día, B=mes). Si B > 12, intentar `mm/dd` (A=mes, B=día). Si ambos fallan → omitir fila.
2. Evaluar si la fecha resultante cae dentro del rango semanal objetivo.

Este método es determinístico, no depende del orden cronológico, y corrige el 100% de los seriales misread.

**Contar** filas (excluyendo encabezado) donde la fecha caiga dentro del rango:
- Para semana actual: `semana_actual_lunes` ≤ Fecha ≤ `semana_actual_hasta`
- Para semana anterior (si aplica): `semana_anterior_lunes` ≤ Fecha ≤ `semana_anterior_domingo`

**4d. Escribir Leads_Calientes (col J) y Leads_Tibios (col K) en REGISTRO_SEMANAL**

Para cada grupo `Codigo_Producto` + `Semana_ISO` en `REGISTRO_SEMANAL` (solo filas con Canal = "Meta"), escribir el conteo en las columnas J y K **únicamente en la primera fila del grupo** (la de menor número de fila); en todas las demás filas del grupo escribir `0`. Esto evita duplicar leads cuando un código tiene varias variantes de campaña — la suma de las columnas J y K debe dar el total real.

Usa Composio `EXCEL_UPDATE_RANGE`:
```
item_id: env ONEDRIVE_ITEM_ID
drive_id: env ONEDRIVE_DRIVE_ID
worksheet_id: "REGISTRO_SEMANAL"
address: "J{n}:K{n}"
values: [[leads_calientes, leads_tibios]]
```

Escribir de forma secuencial, una fila a la vez.

**4e. Avisos de cobertura de clasificación (OBLIGATORIO)**

Un conteo de 0 puede significar dos cosas muy distintas: "esta semana no hubo leads calientes" o "nadie clasificó los leads de este programa". La rutina venía reportando ambas como 0 sin distinguirlas, y por eso el Intake 2 perdió el registro sin que nadie lo notara.

Para cada código procesado, clasificar y reportar:

| Situación | Detección | Significado |
|---|---|---|
| `SIN ARCHIVO` | ningún archivo en MKT CODE empieza con el código | Nunca se creó el archivo de seguimiento. Todos sus leads están sin clasificar |
| `SIN HOJAS` | el archivo existe pero no tiene pestañas `CALIENTE`/`TIBIO` (ej. solo una hoja `Leads`) | Nunca se le corrió el clasificador |
| `HOJAS VACÍAS` | las pestañas existen pero el used range es solo el encabezado (1 fila) | El clasificador se preparó pero nunca se ejecutó |
| `ESTANCADO` | `lastModifiedDateTime` del archivo es anterior al lunes de la semana procesada | Se dejó de clasificar; los ceros no son reales |
| `OK` | hojas con datos y modificado dentro de la semana | Conteo confiable |

Comprobar las pestañas con `EXCEL_LIST_WORKSHEETS` antes de leerlas: si se pide una hoja que no existe, Graph devuelve `ItemNotFound`, que es indistinguible de un archivo borrado y confunde el diagnóstico.

Imprimir siempre un bloque como este, aunque todo esté OK:
```
COBERTURA DE CLASIFICACIÓN <semana>:
  OK:            N códigos
  SIN ARCHIVO:   listar código + leads de Meta que quedan sin clasificar
  SIN HOJAS:     listar
  HOJAS VACÍAS:  listar
  ESTANCADO:     listar código + fecha de última modificación
```

Junto a cada código con problema, mostrar **cuántos leads trajo de Meta en el periodo**: eso convierte el aviso en una cifra accionable ("PE.EI.33-26.1: 35 leads sin clasificar") en vez de una nota técnica.

---

### PASO 5 — Verificación final anti-corrimiento

Al terminar PASO 4, releer el used range completo de `REGISTRO_SEMANAL` y validar:

1. **Sin filas fantasma**: ninguna fila con valores en G:K pero con A-F vacíos. Si existe, limpiarla con `EXCEL_CLEAR_RANGE` (`applyTo: "Contents"`, rango `A{n}:L{n}`).
2. **Correspondencia campaña↔datos**: para cada fila de la semana actual (y W-1 si se procesó), el `ID_Campana` (col E) debe corresponder a los valores G/H/I escritos según los datos de Meta consultados en esta ejecución. Comparar contra el mapa en memoria `{id_campana: (spend, leads, conversaciones)}`. Si alguna fila no coincide, reescribir `G{n}:I{n}` con el valor correcto.
3. **Coherencia de tipo**: campañas `OUTCOME_LEADS` deben tener I=0; campañas `OUTCOME_ENGAGEMENT` deben tener H=0. Si está invertido, corregir. **Además**: campañas `OUTCOME_ENGAGEMENT` (nombre contiene "Whatsapp") con spend > 0 DEBEN tener I > 0 (conversaciones). Si I=0 para una campaña WhatsApp con spend, es un error — la consulta con breakdown no se ejecutó o no se procesó correctamente. Reejecutar paso 2a-bis/3a-bis y corregir.
3b. **Columnas J y K sin corrimiento**: para cada semana procesada, verificar contra los mapas en memoria `{codigo: leads_calientes}` y `{codigo: leads_tibios}` que: (a) la primera fila de cada grupo código+semana tiene exactamente los conteos calculados en PASO 4, (b) las demás filas del grupo tienen 0, y (c) ningún valor de J o K quedó en la fila de un código distinto (corrimiento). Antes de escribir cada `J{n}:K{n}`, releer `D{n}` y confirmar que el código coincide — igual que la regla 10 para G:I. Si hay discrepancia, reescribir las columnas J:K de esa semana.
4. **IDs consecutivos**: la columna A no debe tener saltos ni duplicados. Reportar (no corregir) si se detecta anomalía.
5. **CPL en columna correcta**: verificar que la columna L de cada fila procesada contenga la fórmula CPL (`=IF((H{n}+I{n})=0,"-",G{n}/(H{n}+I{n}))`), y que la columna K contenga un entero (leads tibios), NO una fórmula. Si K contiene una fórmula CPL y L está vacío, la fila se insertó con la estructura de columnas vieja — mover: escribir K=0 y L=fórmula CPL.

Si tras una corrección la validación sigue fallando, detenerse con mensaje de error claro y exit ≠ 0 (no seguir escribiendo a ciegas).

---

## Estructura de filas

Cada fila nueva a insertar tiene 14 valores (columnas A-N):

| Col | Campo | Valor |
|-----|-------|-------|
| A | `ID_Registro` | `ultimo_id + 1` (entero auto-incremental) |
| B | `Fecha_Semana` | serial Excel del lunes de la semana (entero) |
| C | `Semana_ISO` | `"2026-W28"` (string) |
| D | `Codigo_Producto` | código extraído con regex del nombre de campaña |
| E | `ID_Campana` | nombre completo de la campaña en Meta |
| F | `Canal` | `"Meta"` |
| G | `Inversion_USD` | spend (float, 2 decimales) |
| H | `Leads` | leads (entero, 0 para campañas WhatsApp) |
| I | `Conversaciones` | conversaciones iniciadas (entero, 0 para campañas de leads) |
| J | `Leads_Calientes` | `0` al insertar (se actualiza en PASO 4) |
| K | `Leads_Tibios` | `0` al insertar (se actualiza en PASO 4) |
| L | `CPL_USD` | fórmula `=IF((H{n}+I{n})=0,"-",G{n}/(H{n}+I{n}))` donde n = fila Excel |
| M | `Intake` | **fórmula**, nunca texto: `=IFERROR(INDEX('INTAKE 2026'!$B$4:$B$200,MATCH($D{n},'INTAKE 2026'!$C$4:$C$200,0)),"")` |
| N | `Nombre_Programa` | **fórmula**, nunca texto: `=IFERROR(INDEX('INTAKE 2026'!$D$4:$D$200,MATCH($D{n},'INTAKE 2026'!$C$4:$C$200,0)),"")` |

**M y N son obligatorias y SIEMPRE se escriben, al insertar y al actualizar.** El rango de escritura de una fila nueva es `A{n}:N{n}` — 14 columnas. Escribir `A:L` deja M y N vacías, y entonces todos los paneles del `DASHBOARD INTAKE 2` que filtran por `M = "Intake 2"` descartan la fila en silencio: la semana aparece en cero aunque los datos estén en el registro.

Esto ya ocurrió dos semanas seguidas: W37 (19 de 25 filas sin M, corregido el 2026-09-11) y W38 (23 filas, corregido el 2026-09-18). Ambas veces se reparó a mano sin tocar la rutina, y volvió a pasar.

Por qué fórmula y no texto: la hoja `INTAKE 2026` es el catálogo maestro (columna B = intake, C = código, D = nombre). Con la fórmula, dar de alta un programa ahí basta para que todas sus filas se clasifiquen solas; la rutina no necesita conocer el catálogo. El rango llega a la fila 200 a propósito: el rango original terminaba en la 29 y el catálogo ya tiene códigos en la 30 (`PE.EI.34-26.2` quedaba fuera). Devuelve `""` y no `"No encontrado"` para que un código aún sin dar de alta no ensucie los filtros.

Si al escribir una fila el código no está en `INTAKE 2026`, la fórmula devolverá `""`: reportarlo en el resumen final como "código sin dar de alta en INTAKE 2026".

Escribir con Composio `EXCEL_UPDATE_RANGE`:
- `worksheet_id`: `REGISTRO_SEMANAL`
- `address`: `A{primera_fila_vacia}:N{primera_fila_vacia}` (una fila a la vez)

Al actualizar filas existentes (PASO 2 Caso B y PASO 3):
- Spend/leads/conv → `G{n}:I{n}`
- Leads_Calientes → `J{n}`, Leads_Tibios → `K{n}` (en PASO 4, después de procesar MKT CODE)
- **No tocar L (CPL, es fórmula)**

---


## Variables de entorno requeridas

| Variable | Descripción |
|---|---|
| `META_ACCESS_TOKEN` | Token de acceso largo de Meta Ads (solo si el Meta MCP lo requiere) |
| `META_AD_ACCOUNT_ID` | ID numérico de la cuenta publicitaria (solo los dígitos, sin `act_`) |
| `ONEDRIVE_ITEM_ID` | `01EJAD6P26K7XV47MKCNA2NUMJMCTYHSQA` |
| `ONEDRIVE_DRIVE_ID` | `b!1U4iaBDsVk2MoPFpxkox96PVSF7eIfZPn1_TQQsa_Rux7EmX3JabSbzUbWh20VZS` |
| `MKT_CODE_DRIVE_ID` | `b!JDzJFFxcZ0i_gjgWsk8jlYN--xjLhhVPsUr6Gl5Fm-d9EsShc690ToWO57RPZpCl` |
| `MKT_CODE_FOLDER_ID` | `01XEGP2KYCIGWDO6XTTVBZDOF2IIBDMQ5S` |
| `VENTAS_DRIVE_ID` | `b!fzUTGJaceE2y6ZsHwbhgVyxuBRCiSZtAsxy9cziifxfnoz3mV6j-TbvlIQ7dFiRT` (josecardenas, SOLO LECTURA) |

## Catálogo de códigos → Intake + Nombre

**Las columnas M y N ya NO se rellenan con este catálogo** — se escriben como fórmulas contra la hoja `INTAKE 2026` (ver "Estructura de filas"). Esta tabla queda solo como referencia para leer el registro y para saber qué códigos esperar. Si discrepa de `INTAKE 2026`, manda `INTAKE 2026`.

Nunca asignar `"Intake 1"` por defecto a un código desconocido: mete programas en el intake equivocado sin que se note.

### Intake 1 (12 códigos)

| Código | Nombre_Programa |
|--------|-----------------|
| PE.EI.08-26.2 | Automatización IEC 61850 |
| PE.EI.10-26.2 | Protección de Sistemas Eléctricos de Potencia |
| CE.EI.06-26.1 | Auditorías Técnicas Fotovoltaicas |
| PE.EI.07-26.2 | Contratos y Negociaciones en el Mercado Eléctrico Mexicano |
| PE.EI.27-26.1 | Transitorios Electromagnéticos con ATP-EMTP |
| PE.EI.14-26.2 | Sistemas Solares Fotovoltaicos |
| PE.EI.37-26.1 | Energy Data Analytics |
| PE.EI.05-26.2 | Diseño de Instalaciones Eléctricas de Baja Tensión |
| PE.EI.20-26.2 | Diseño de Líneas de Transmisión de Alta y Extra Alta Tensión |
| PE.EI.06-26.4 | Sistemas de Almacenamiento de Energía en Baterías (BESS) |
| CE.EI.02-26.3 | Análisis Financiero de Energías Renovables |
| PE.EI.34-26.1 | Sistemas de Almacenamiento de Energía en Baterías (BESS) - Colombia |

### Intake 2 (11 códigos)

| Código | Nombre_Programa |
|--------|-----------------|
| PE.EI.23-26.2 | Compensadores FACTS en Sistemas de Transmisión |
| PE.EI.26-26.1 | Estudios de Conexión de Generación Renovable a la RED |
| PE.EI.31-26.1 | Asset Management y Operación Estratégica de Sistemas Híbridos |
| PE.EI.33-26.1 | Desarrollo y Gestión de Proyectos de Energías Renovables |
| PE.EI.10-26.3 | Protección de Sistemas Eléctricos de Potencia |
| PE.EI.36-26.1 | Sistemas de Almacenamiento de Energía en Baterías (BESS) - CHILE |
| PE.EI.38-26.1 | Nuevas Tecnologías en Sistemas de Distribución |
| CE.EI.09-26.1 | Regulación 4.0 para la Distribución Eléctrica |
| CE.EI.03-26.2 | Análisis Financiero e Inversiones en SAE BESS |
| PE.EI.06-26.5 | Sistemas de Almacenamiento de Energía en Baterías (BESS) |
| PE.EI.39-26.1 | Servicios Complementarios en el Perú |

---

### PASO 5-bis — Reconciliación semanal completa contra Meta (OBLIGATORIO)

Ejecutar SIEMPRE, en toda corrida, para las dos semanas procesadas (actual y anterior). Este paso existe porque en la corrida del 2026-09-04 se detectaron 2 campañas con gasto que nunca se registraron, 1 código escrito mal, 4 filas duplicadas por renombre y 33 filas fantasma.

**5b-1. Construir el set de referencia desde Meta**

Para cada semana procesada, tomar la respuesta de `ads_get_ad_entities` y construir:
```
meta_semana = { name: (amount_spent, leads, conversaciones, objective, results_indicator) }
```
incluyendo **todas** las campañas con `amount_spent > 0`, incluso las que se van a omitir por no tener código.

**5b-2. Comparar contra el Excel en las dos direcciones**

Leer todas las filas de esa semana con Canal = "Meta" y comparar:

| Caso | Detección | Acción |
|---|---|---|
| **Gasto no registrado** | Campaña en `meta_semana` con código válido que NO tiene fila en el Excel | Insertar la fila faltante |
| **Fila obsoleta / renombre** | Fila en el Excel cuyo `ID_Campana` no existe en `meta_semana` | Verificar si es un renombre (mismo código, otra fila con nombre nuevo). Si lo es → limpiar la fila obsoleta con `EXCEL_CLEAR_RANGE`. Si no → reportar, no borrar |
| **Gasto desactualizado** | `G/H/I` del Excel ≠ valores de `meta_semana` | Reescribir `G{n}:I{n}` con los valores de Meta |
| **Código mal extraído** | El código de la col D no coincide con el regex aplicado al `ID_Campana` de la col E | Corregir la col D, y recalcular M y N desde el catálogo |
| **Fila fantasma** | Col D vacía pero G/H/I o L/M con contenido | Limpiar `A{n}:N{n}` con `EXCEL_CLEAR_RANGE` (`applyTo: "Contents"`) |
| **Duplicado exacto** | Dos filas con mismo código + semana + `ID_Campana` | Conservar la de mayor gasto (la más reciente), limpiar la otra |

**5b-2-bis. Prueba de campañas huérfanas sobre TODO el histórico (no solo la semana en curso)**

Un renombre puede haber ocurrido semanas atrás y quedar sin detectar. Por eso esta prueba corre sobre el registro completo, no sobre la semana procesada:

1. Consultar Meta una vez con el rango completo que cubre el registro:
   ```
   level: "campaign"
   fields: ["id","name","objective","amount_spent","lead","results"]
   filtering: [{"field":"campaign.amount_spent","operator":"GREATER_THAN","value":["0"]}]
   time_range: {"since": <lunes de la primera semana del registro>, "until": <hoy>}
   limit: 200
   ```
2. Construir `nombres_meta = set(name)` con **todas** las campañas devueltas.
3. Recorrer el registro y agrupar por `ID_Campana` (col E). **Todo `ID_Campana` que no esté en `nombres_meta` es una campaña huérfana**: el nombre ya no existe en Meta, lo que solo puede significar que la campaña fue renombrada.
4. Para cada huérfana, buscar en el registro las filas del mismo `Codigo_Producto` cuyo nombre **sí** exista en Meta y sea el más parecido (mismo prefijo). Ese es el nombre nuevo. Entonces:
   - Si hay una fila con el nombre nuevo **en la misma semana** → la fila huérfana es un duplicado: limpiarla con `EXCEL_CLEAR_RANGE`. Antes de limpiar, comprobar si carga `J`/`K` (leads calientes/tibios); si la fila que sobrevive los tiene en 0 y la huérfana no, trasladar los valores.
   - Si NO hay fila con el nombre nuevo en esa semana → la fila es legítima, solo tiene el nombre viejo: actualizar `E{n}` al nombre actual de Meta.
5. Validar el resultado: para cada campaña, el gasto acumulado en el registro debe coincidir con el de Meta con tolerancia de **±$0.50** (redondeo de céntimos por semana). Si la diferencia supera eso, pedir el desglose semanal a Meta con `object_ids: [<id>]` y `time_increment: "7"` y corregir la semana que no cuadre.

**Por qué importa**: en la corrida del 2026-09-04, `PE.EI.37-26.1 Energy Data 6%` fue renombrada a `Energy Data 6%.` (se le agregó un punto). La rutina había corrido antes y después del renombre en la misma semana W34, generando dos filas: $37.64 con el nombre viejo y $39.35 con el nuevo. Eso infló el código en $37.64 y duplicó también sus leads calientes/tibios. El registro reportaba $653.45 cuando Meta decía $615.90.

**Aviso**: el nombre de campaña NO es una clave estable. Mientras el registro no guarde el `id` numérico de la campaña, esta prueba de huérfanas es la única defensa contra los renombres — es obligatoria en cada corrida, no opcional.

**5b-2-ter. Detección de gasto con dígito perdido**

Comparar el gasto de cada fila contra el desglose semanal de Meta. Un gasto que difiere en exactamente un factor de 10 o 100 (ej. registro `$0.18` vs Meta `$18.01`) indica un valor truncado al escribir. Señal de alarma barata: una fila con `Leads` alto y `Inversion_USD` cerca de cero (CPL absurdamente bajo, < $0.10) casi siempre es esto. Verificar contra Meta y corregir.

**5b-3. Validación estructural de toda la hoja**

Recorrer todas las filas y verificar:
- Col B es **numérica** (serial) y col C es **string** `"2026-Wnn"`. Si están invertidas → es una fila escrita por una versión vieja; corregir o limpiar.
- Col F (`Canal`) == `"Meta"`. Valores como `"Meta Ads"` o `"WhatsApp"` provienen de versiones antiguas de la rutina y suelen ser duplicados parciales — verificar contra Meta antes de limpiar.
- Col M (`Intake`) y col N (`Nombre_Programa`) no vacías. Si están vacías, completarlas desde el catálogo.
- Col L es **fórmula** CPL con el formato actual `=IF((H{n}+I{n})=0,"-",G{n}/(H{n}+I{n}))`. La fórmula vieja `=IF(H{n}=0,"",G{n}/H{n})` indica fila de versión antigua.
- No hay dos filas con el mismo `ID_Registro`.

**5b-4. Reporte obligatorio**

Imprimir SIEMPRE, aunque no haya hallazgos:
```
RECONCILIACIÓN <semana>:
  campañas en Meta con gasto:      N
  campañas omitidas (sin código):  N  → listar nombres
  filas insertadas (gasto perdido): N → listar
  filas actualizadas:               N
  filas limpiadas (fantasma/dup):   N → listar con motivo
  códigos corregidos:               N → listar viejo → nuevo
```

**5b-4-bis. Comprobación final de M/N (bloqueante)**

Justo antes de terminar, releer `REGISTRO_SEMANAL` completo y contar las filas con código (col D) cuya columna M esté **vacía**. Si hay alguna:
1. Escribir en `M{n}:N{n}` las fórmulas de la tabla "Estructura de filas".
2. Releer y volver a contar.
3. Si sigue habiendo filas con M vacía, terminar con exit ≠ 0 y listarlas.

Revisar **todo el registro**, no solo las semanas procesadas: en W38 también había una fila de W37 (la 274) sin M.

Esta comprobación existe porque las validaciones anteriores no bastaron: la regla de escribir `A:N` ya estaba documentada y aun así la rutina escribió `A:L` dos semanas seguidas. Una fila sin M no produce ningún error visible; solo hace que el dashboard muestre la semana en cero.

Si tras las correcciones la validación 5b-3 sigue fallando, detenerse con exit ≠ 0 y no continuar al PASO 6.

---

### PASO 6 — Ventas Intake 2 desde Cuadro de ventas diarias → TABLERO MKT-VENTAS

Ejecutar **después** de PASO 5. Este paso lee las ventas acumuladas de cada programa Intake 2 desde el archivo externo "Cuadro de ventas diarias" (propiedad de josecardenas@inelinc.com) y las escribe en la columna L ("Ventas (u)") de la hoja `TABLERO MKT-VENTAS` del registro.

**⚠️ REGLA ABSOLUTA: el archivo "Cuadro de ventas diarias" es SOLO LECTURA. Nunca escribir, modificar ni subir nada a ese archivo. Solo leer.**

**6a. Leer el Cuadro de ventas diarias — por Composio, NO por `read_resource`**

```
tool:         EXCEL_GET_RANGE   (Composio)
item_id:      01EOWQSYVLOKNQ4QMLEFF3M6IHN7LR7O46
drive_id:     b!fzUTGJaceE2y6ZsHwbhgVyxuBRCiSZtAsxy9cziifxfnoz3mV6j-TbvlIQ7dFiRT
worksheet_id: "V.PROGRAMAS 2026"
address:      "A1:H600"      ← ampliar si la hoja crece (ver abajo)
```

Solo se necesitan las columnas A–H; el resto de la hoja no aporta nada al conteo de ventas. Ejecutar la llamada **dentro del sandbox** (`COMPOSIO_REMOTE_WORKBENCH` con `run_composio_tool`) para que las ~600 filas no pasen por el contexto.

Si `item_id` deja de servir, reubicarlo con `ONE_DRIVE_SEARCH_ITEMS` pasando `drive_id` y `q: "Cuadro de ventas diarias"` — **el parámetro es `q`, no `query`**.

**⚠️ `read_resource` del MCP de Microsoft YA NO SIRVE para este archivo.** Hasta el 2026-09-08 funcionaba. El 2026-09-10 le agregaron a `V.PROGRAMAS 2026` cuatro columnas de fórmulas (T/U/V/W, una fórmula por fila × 599 filas ≈ 2.400 líneas) que consumen el presupuesto del volcado **antes** de llegar a los datos: el bloque `INTAKE 2` desaparece de la extracción aunque esté en la hoja. El fallo es silencioso — parece que la sección no existe. No volver a `read_resource` aunque parezca más simple.

**Ajustar el rango a la hoja.** La hoja crece: 471 filas el 04/09, 599 el 10/09. Antes de leer, confirmar el tamaño con `EXCEL_GET_WORKSHEET_USED_RANGE` (o pedir un rango holgado, ej. `A1:H900`) y verificar que la última fila devuelta esté vacía. Si el rango termina justo en datos, se está cortando el final de la hoja — ampliarlo y releer.

**Otra trampa:** la hoja oculta `V. ASINCRONOS 2026` es un worksheet distinto, así que leyendo por rango ya no se mezcla. Pero no leer nunca la hoja entera sin acotar el `worksheet_id`.

**6b. Parsear ventas por programa Intake 2**

En el array `values` que devuelve `EXCEL_GET_RANGE`, localizar la fila que contenga una celda cuyo texto sea exactamente **`INTAKE 2`**. El bloque va desde la fila siguiente hasta el final de los datos. Todo lo anterior es Intake 1 o intakes previos — no mezclar.

**Nunca fijar el número de fila del marcador.** Se mueve cada vez que se agregan programas al Intake 1 (fila 473 el 10/09/2026). Buscarlo siempre.

Cada programa ocupa una fila principal seguida de filas `SEM n:` con el desglose semanal. **Usar siempre la fila principal**, nunca la suma de las SEM: varias filas SEM traen rangos de fecha mal escritos (ej. `SEM 8: 12/08 al 17/10`, con el mes de inicio equivocado) y no son fiables.

Índices dentro de cada fila del array (columnas A–H):

| Índice | Columna | Campo |
|---|---|---|
| 0 | A | número correlativo del programa |
| 1 | B | INICIO |
| 2 | C | FIN |
| 3 | D | **PROGRAMAS** (nombre) |
| 4 | E | RESPONSABLE |
| 5 | F | **CODIGO** |
| 6 | G | COMENTARIOS |
| 7 | H | **VENTAS** ← el acumulado que se registra |

**Criterio de filtrado**: solo filas cuyo índice 0 sea **numérico** (`isinstance(row[0], (int, float))`). Las `SEM n:` traen ese campo vacío. Ojo: leyendo por rango los vacíos llegan como `''`, no como `None`.

**Detección de código duplicado (OBLIGATORIO)**

El campo CODIGO se escribe a mano y se equivoca. Verificado el 2026-09-08: el programa 56 se llama **"BESS V"** pero su CODIGO dice `CE.EI.09-26.1`, que es Regulación 4.0 y ya está en el programa 53. Sus 3 ventas corresponden a `PE.EI.06-26.5`.

Por eso, tras construir el mapa:
1. Si un código aparece **dos veces** dentro del bloque INTAKE 2 → hay un error de tipeo. NO sumar los dos valores.
2. Resolver por el campo PROGRAMAS (índice 3), que sí es fiable, contra el nombre del catálogo. `BESS V` = quinto BESS = `PE.EI.06-26.5`; `BESS CHILE` = `PE.EI.36-26.1`; etc.
3. Si un código del catálogo queda sin fila **y** hay un duplicado sin resolver, lo más probable es que el duplicado le pertenezca.
4. Reportar siempre el duplicado en el resumen final, con el código, el nombre del programa y la atribución que se aplicó. Es una inferencia sobre un archivo ajeno que nadie más está validando.

**Cobertura del catálogo**: `routine-amg-semanal/programas.json` lista **13** programas de Intake 2, pero el TABLERO solo tiene **11** filas (9-19). Faltan `PE.EI.34-26.2` y `SM.EI.03-26.1`, que no tienen dónde escribirse. Si aparecen con ventas en el cuadro, reportarlo — no inventar filas en el tablero sin avisar.

**6c. Escribir ventas en TABLERO MKT-VENTAS**

La hoja `TABLERO MKT-VENTAS` del registro tiene:
- Worksheet ID: `{6529AB3D-6367-4338-AC95-EFBADF3FC760}`
- Encabezado en fila 8
- Datos de programas desde la fila 9
- Fila `TOTAL` inmediatamente después del último programa
- Columna A = código del programa
- Columna L = Ventas (u) — **esta es la columna destino**

**El número de programas NO es fijo.** Eran 11 (filas 9-19, TOTAL en 20) hasta que el usuario agregó el Summit el 2026-09-09; ahora son 12 (filas 9-20, TOTAL en **21**). Va a volver a cambiar.

Por eso: **leer siempre la columna A hacia abajo desde la fila 9** hasta encontrar la primera celda vacía, y tratar esa fila como la de TOTAL. Nunca cablear `L9:L19` ni asumir dónde está el total.

Para cada fila de programa, tomar el código de la columna A, buscarlo en el mapa `{codigo: ventas_totales}` y escribir el valor en su celda L. **Si un código del tablero no está en el mapa, abortar el paso y reportarlo** — escribir 0 por defecto enmascararía un desalineamiento.

Usa Composio `EXCEL_UPDATE_RANGE`:
```
item_id: env ONEDRIVE_ITEM_ID
drive_id: env ONEDRIVE_DRIVE_ID
worksheet_id: "TABLERO MKT-VENTAS"
address: "L{n}"    (n = fila del programa, 9 a 19)
values: [[ventas]]
```

Escribir de forma secuencial, una celda a la vez, para evitar corrimiento. Antes de escribir `L{n}`, verificar que `A{n}` coincide con el código esperado.

**La columna L está vacía, sin fórmulas.** El texto de ayuda del propio tablero dice que L "lee VENTAS_INPUT y solo suma filas con Estado = Confirmada", pero esa hoja nunca se creó y las celdas no tienen fórmula (verificado 2026-09-08). Escribir valores literales es correcto y no pisa nada.

**6d. Total y verificación**

La fila TOTAL necesita `=SUM(L9:L{ultima_fila_programa})`. Comprobar si ya la tiene: al 2026-09-10 la fila 21 sí trae `=SUM(L9:L20)` porque se escribió en la corrida anterior y el usuario la extendió al agregar el Summit. Si el número de programas cambió, **corregir el rango de la fórmula**, no solo los valores.

(Una versión anterior de este documento decía que la fila TOTAL ya tenía fórmula de suma y que no se tocara: era falso, estaba vacía.)

Después de escribir, releer desde `A9` hasta la fila TOTAL y validar que cada fila tiene el valor correcto según el mapa y que el total cuadra con la suma de las ventas leídas del cuadro. **Ojo:** al releer inmediatamente después de escribir, la celda del total puede devolver todavía el valor viejo porque Excel no ha recalculado. Releer una segunda vez antes de dar por fallida la validación.

**6e. Cuándo no hay nada que registrar**

Que el bloque `INTAKE 2` no exista en el cuadro **no es un error**. Los programas se dan de alta cuando arranca su fase comercial, no antes. Sucedió el 2026-09-04: la hoja terminaba en el programa 44 y la sección apareció recién unos días después.

Si el marcador `INTAKE 2` no aparece:
- No escribir nada en la columna L (no rellenar con ceros: borraría datos previos).
- Reportar `"Cuadro de ventas: sección INTAKE 2 no encontrada — 0 programas leídos"` y terminar el paso con éxito.

**Antes de concluir que la sección no existe, descartar que sea un problema de lectura.** Este diagnóstico ya se dio por bueno una vez estando equivocado: el 2026-09-10 el marcador faltaba solo porque `read_resource` truncaba, y la sección llevaba días creada con datos.

Comprobaciones obligatorias antes de reportar "no existe":
1. ¿La respuesta trae realmente las ~600 filas, o vino recortada? Contar `len(values)`.
2. ¿Aparece el marcador `INTAKE 1`? Si tampoco está, **no** es que falte la sección: la lectura falló. Ampliar el rango y reintentar.
3. ¿La última fila del rango tiene datos? Entonces el rango se quedó corto — ampliarlo.

Solo si `INTAKE 1` aparece, el rango cubre toda la hoja y aun así no hay `INTAKE 2`, la sección realmente no se ha creado. Eso pasó el 2026-09-04 y es legítimo: los programas se dan de alta cuando arranca su fase comercial.

---

## Resumen final

Imprime al terminar:
- Semanas procesadas (actual + anterior si aplicó backfill)
- Filas insertadas vs actualizadas
- Campañas omitidas (sin código)
- Leads calientes y tibios por código (ej. `PE.EI.34-26.1: 5 calientes, 9 tibios`); códigos sin archivo en MKT CODE: listar
- Ventas Intake 2 escritas en TABLERO MKT-VENTAS: listar código → ventas por cada programa, más el total. Añadir:
  - códigos duplicados en el cuadro y cómo se resolvieron (ver 6b)
  - programas del catálogo sin fila en el tablero
  - si la sección INTAKE 2 aún no existía

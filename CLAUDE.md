# Routine: Registro semanal Meta Ads → Excel OneDrive

## Objetivo
Al ejecutarse en cualquier día, registrar en la hoja `REGISTRO_SEMANAL` del Excel de OneDrive la inversión, leads y conversaciones de Meta Ads para la semana actual (lunes → hoy) y, si falta, también la semana anterior completa (lunes → domingo). Una fila por campaña por semana. Campañas de leads (`OUTCOME_LEADS`) registran leads; campañas de WhatsApp (`OUTCOME_ENGAGEMENT`) registran conversaciones iniciadas.

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
10. **Respetar filas de otros canales**: al leer `filas_existentes`, ignorar para todo propósito de actualización las filas donde columna F (`Canal`) sea distinto de `"Meta"` (ej. `"WhatsApp Masivo"`, `"LinkedIn"`). Esas filas son de entrada manual — no tocarlas, no sobreescribirlas, no usarlas para calcular `ultimo_id` ni `primera_fila_vacia`.

## Asignación MCP vs script

| Operación | Método |
|---|---|
| Consultar campañas, spend y leads de Meta Ads | Meta MCP (`ads_get_ad_entities`) |
| Leer hoja `REGISTRO_SEMANAL` | Composio `EXCEL_GET_WORKSHEET_USED_RANGE` |
| Actualizar celdas de fila existente | Composio `EXCEL_UPDATE_RANGE` |
| Insertar nuevas filas | Composio `EXCEL_UPDATE_RANGE` |

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

**2a. Consultar Meta Ads para la semana anterior completa**

Usa Meta MCP `ads_get_ad_entities`:
```
ad_account_id: env META_AD_ACCOUNT_ID
level: "campaign"
fields: ["id", "name", "objective", "amount_spent", "lead", "results"]
filtering: [{"field": "campaign.amount_spent", "operator": "GREATER_THAN", "value": ["0"]}]
time_range: {"since": "semana_anterior_lunes", "until": "semana_anterior_domingo"}
```

**2b. Filtrar y procesar campañas de la semana anterior**

Por cada campaña en los resultados:
- Si `amount_spent` = 0 y `lead` = 0 y `results` = 0 → omitir
- Extraer `codigo` con regex `(PE\.EI\.\d+-\d+\.\d+|CE\.EI\.\d+-\d+\.\d+)` del campo `name`
- Si no hay coincidencia con PE.EI o CE.EI → omitir silenciosamente
- `id_campana` = `name` completo tal como viene de Meta
- Si `objective` == `OUTCOME_ENGAGEMENT` → campaña WhatsApp: `leads = 0`, `conversaciones` = valor numérico de `results`
- Si `objective` == `OUTCOME_LEADS` → campaña de leads: `leads` = valor de `lead`, `conversaciones = 0`

**2c. Escribir filas de la semana anterior**

Para cada campaña válida, insertar una fila nueva (no hay duplicados posibles porque la semana anterior no estaba en el Excel).

Ver estructura de columnas en la sección **Estructura de filas** más abajo.

Actualizar `ultimo_id` y `primera_fila_vacia` a medida que se insertan filas.

---

### PASO 3 — Procesar semana actual

**3a. Consultar Meta Ads para la semana actual**

Usa Meta MCP `ads_get_ad_entities`:
```
ad_account_id: env META_AD_ACCOUNT_ID
level: "campaign"
fields: ["id", "name", "objective", "amount_spent", "lead", "results"]
filtering: [{"field": "campaign.amount_spent", "operator": "GREATER_THAN", "value": ["0"]}]
time_range: {"since": "semana_actual_lunes", "until": "semana_actual_hasta"}
```

**3b. Filtrar campañas**

Por cada campaña:
- Si `amount_spent` = 0 y `lead` = 0 y `results` = 0 → omitir
- Extraer `codigo` con regex del campo `name` → si no tiene PE.EI o CE.EI → omitir
- `id_campana` = `name` completo
- Si `objective` == `OUTCOME_ENGAGEMENT` → campaña WhatsApp: `leads = 0`, `conversaciones` = valor numérico de `results`
- Si `objective` == `OUTCOME_LEADS` → campaña de leads: `leads` = valor de `lead`, `conversaciones = 0`

**3c. Decidir: ¿insertar o actualizar?**

Para cada campaña válida, buscar en `filas_existentes` si existe una fila donde:
- Columna C (`Semana_ISO`) == `semana_actual_iso`
- Columna E (`ID_Campana`) == `campaign_name` (exacto)

**Si la fila EXISTE → ACTUALIZAR**
Usa Composio `EXCEL_UPDATE_RANGE` en las celdas G, H e I de esa fila:
- `address`: `G{n}:I{n}` (n = número de fila Excel de la coincidencia)
- `values`: [[spend, leads, conversaciones]]

**Si la fila NO EXISTE → INSERTAR**
Insertar nueva fila (ver estructura abajo) en `primera_fila_vacia`.
Incrementar `ultimo_id` y `primera_fila_vacia`.

---

## Estructura de filas

Cada fila nueva a insertar tiene 11 valores (columnas A-K):

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
| J | `CPL_USD` | fórmula `=IF((H{n}+I{n})=0,"-",G{n}/(H{n}+I{n}))` donde n = fila Excel |
| K | `Comentario` | `""` |

Escribir con Composio `EXCEL_UPDATE_RANGE`:
- `worksheet_id`: `REGISTRO_SEMANAL`
- `address`: `A{primera_fila_vacia}:K{primera_fila_vacia}` (una fila a la vez) o en bloque si son varias filas consecutivas

---


## Variables de entorno requeridas

| Variable | Descripción |
|---|---|
| `META_ACCESS_TOKEN` | Token de acceso largo de Meta Ads (solo si el Meta MCP lo requiere) |
| `META_AD_ACCOUNT_ID` | ID numérico de la cuenta publicitaria (solo los dígitos, sin `act_`) |
| `ONEDRIVE_ITEM_ID` | `01EJAD6P26K7XV47MKCNA2NUMJMCTYHSQA` |
| `ONEDRIVE_DRIVE_ID` | `b!1U4iaBDsVk2MoPFpxkox96PVSF7eIfZPn1_TQQsa_Rux7EmX3JabSbzUbWh20VZS` |

## Resumen final

Imprime al terminar:
- Semanas procesadas (actual + anterior si aplicó backfill)
- Filas insertadas vs actualizadas
- Campañas omitidas (sin código)

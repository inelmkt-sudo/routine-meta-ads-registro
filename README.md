# routine-meta-ads-registro

Registra automáticamente la inversión y leads de Meta Ads por campaña en la hoja `REGISTRO_SEMANAL` del Excel de OneDrive de INEL. Una fila por campaña por semana.

## Qué hace

1. Consulta Meta Ads (MCP) con filtro `amount_spent > 0` y obtiene inversión + leads por campaña
2. Filtra solo campañas con código PE.EI o CE.EI — las demás se omiten
3. Lee la hoja `REGISTRO_SEMANAL` vía Composio para verificar estado actual
4. Agrega una fila por campaña (idempotente por `ID_Campana` + `Semana_ISO`: actualiza si ya existe)
5. Si la semana anterior no está registrada, hace backfill automático (solo W-1)
6. El CPL se calcula como fórmula en la celda, no como valor

## Archivo Excel

- **Archivo**: `REGISTRAR PROGRAMAS WORKSHOPS.xlsx`
- **Ruta OneDrive**: `09. Marketing / INSTITUTO / REGISTRO`
- **Hoja**: `REGISTRO_SEMANAL`
- **Drive ID**: `b!1U4iaBDsVk2MoPFpxkox96PVSF7eIfZPn1_TQQsa_Rux7EmX3JabSbzUbWh20VZS`
- **Item ID**: `01EJAD6P26K7XV47MKCNA2NUMJMCTYHSQA`

## Setup

Ver [docs/SETUP.md](docs/SETUP.md) para obtener las credenciales.

## Variables de entorno

```
META_ACCESS_TOKEN
META_AD_ACCOUNT_ID
ONEDRIVE_ITEM_ID
ONEDRIVE_DRIVE_ID
```

## Configuración del Routine

- **Environment**: `bash setup.sh`
- **Network access**: Full
- **Connectors**: Meta Ads MCP + Composio (con conexión Excel activa para natalieaguirre@inelinc.com)
- **Trigger**: Weekly o Daily — hora de tu preferencia
- **Prompt**:

```
Lee CLAUDE.md y ejecuta la automatización descrita ahí. Operas con autonomía
plena: no preguntas, no pides confirmación, no esperas input — nadie puede
contestarte. Tienes libertad total para decidir según las instrucciones y tu
juicio (como --dangerously-skip-permissions). Si algo es genuinamente
imposible, detente con nota clara y exit error. Respeta la asignación MCP vs
script del CLAUDE.md. Reporta resumen al final.
```

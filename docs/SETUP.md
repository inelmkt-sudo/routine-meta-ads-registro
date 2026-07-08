# Guía de configuración

## 1. Meta Ads — Access Token y Ad Account ID

### Access Token
1. Ve a [developers.facebook.com/tools/explorer](https://developers.facebook.com/tools/explorer)
2. Selecciona tu app (o crea una en Meta for Developers)
3. Agrega los permisos: `ads_read`, `read_insights`
4. Haz clic en **Generate Access Token**
5. Para un token que no expire: crea un token de sistema desde **Business Settings → System Users**

### Ad Account ID
1. Ve a [business.facebook.com](https://business.facebook.com)
2. Abre **Administrador de anuncios**
3. El ID aparece en la URL: `act_XXXXXXXXX` — copia solo los números

---

## 2. OneDrive — Item ID y Drive ID

El archivo y los IDs ya están identificados. No necesitas buscarlos a mano.

| Variable | Valor |
|---|---|
| `ONEDRIVE_ITEM_ID` | `01EJAD6P26K7XV47MKCNA2NUMJMCTYHSQA` |
| `ONEDRIVE_DRIVE_ID` | `b!1U4iaBDsVk2MoPFpxkox96PVSF7eIfZPn1_TQQsa_Rux7EmX3JabSbzUbWh20VZS` |

**Archivo**: `REGISTRAR PROGRAMAS WORKSHOPS.xlsx`
**Ruta**: `09. Marketing / INSTITUTO / REGISTRO` en el OneDrive de natalieaguirre@inelinc.com
**Hoja destino**: `REGISTRO_SEMANAL`

---

## 3. Composio — Conexión Excel

El Routine escribe en el Excel vía el conector Composio Excel (sin Power Automate).

La conexión **ya está activa** en Composio para la cuenta `natalieaguirre@inelinc.com` (cuenta `excel_salva-shuff`).

Para configurar el Routine en `claude.ai/code/routines`:
1. En el Routine, agrega **Composio** como conector (MCP remoto)
2. Asegúrate de que la conexión Excel esté vinculada a `natalieaguirre@inelinc.com`
3. No se necesita ninguna App Registration en Azure ni credenciales adicionales


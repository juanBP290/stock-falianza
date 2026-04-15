# Apps Script

Este directorio guarda el backend del modulo de servicios.

## Archivo principal

- `Code.gs`: backend corregido con soporte para:
  - `delete_proforma`
  - `upsert_proforma`
  - lectura de `vehiculos` y `servicios`
  - tokens por propiedades del script

## Paso manual que falta

Desde Codex no tengo acceso directo para pegar codigo dentro del proyecto de Apps Script ni para redeployar tu Web App. Para activarlo:

1. Abre tu proyecto de Google Apps Script.
2. Reemplaza el contenido de tu archivo `.gs` con `apps-script/Code.gs`.
3. Guarda los cambios.
4. Ve a `Deploy` > `Manage deployments`.
5. Edita tu Web App y crea una `New version`.
6. Despliega de nuevo.
7. Si cambia la URL `/exec`, actualizala en la pestana `Config` de la web.

## Script Properties recomendadas

- `TOKEN_WRITE`
- `TOKEN_READ`
- `SPREADSHEET_ID` (solo si el script no esta vinculado a la hoja)

## Nota sobre Google Sheets

La hoja `servicios` debe tener al menos esta columna para eliminar proformas:

- `proforma_num`

Y para guardar mejor los datos del formulario conviene tambien:

- `cliente`
- `doc`
- `telefono`
- `direccion`
- `ano`

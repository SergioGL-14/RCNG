# Mantenimiento

Guia corta de mantenimiento operativo.

## Actualizar Datos

El boton `Actualizar Datos` realiza estas acciones:

1. intenta sincronizar con el recurso compartido configurado;
2. refresca `ComputerNames.sqlite`;
3. vuelve a leer `scripts_db.json` y `apps_db.json`;
4. reconstruye los menus `Scripts` y `Aplicacions`;
5. informa en el panel de logs.

Usarlo cuando:

- se hayan cambiado scripts o iconos;
- se hayan anadido apps externas compartidas;
- se haya actualizado la base de equipos en servidor;
- la copia local parezca desalineada.

## Mantenimiento de equipos

Ruta:

- `Configuracion -> Equipos`

Operaciones disponibles:

- buscar;
- paginar;
- anadir;
- editar;
- eliminar;
- importar CSV;
- guardar;
- replicar la base al recurso compartido.

## Mantenimiento de extensiones

Ruta:

- `Aplicacions -> ExtFinder`

Operaciones disponibles:

- buscar por equipo;
- buscar por extension;
- alta;
- modificacion;
- borrado.

## Mantenimiento del entorno

Ruta:

- `Configuracion -> Entorno global`

Desde ahi se ajustan:

- rutas UNC;
- proxy PAC;
- portales web;
- ubicacion del inventario para WOL.

## Scripts y aplicaciones

### Scripts

- alta desde `Scripts -> Anadir script...`;
- baja o modificacion desde el menu contextual del propio script;
- replicacion opcional al recurso compartido.

### Apps externas

- alta desde `Aplicacions -> Anadir nuevas aplicaciones...`;
- modificacion y borrado desde el menu contextual de la app;
- replicacion opcional de `apps_db.json` e iconos.

## WOL

El boton `Encender` depende del contenido de `app\WOL\`.

Revision recomendada:

- que exista `WOL.vbs`;
- que existan `mc-wol.exe` y `PsExec.exe`;
- que el CSV local tenga formato correcto;
- que la ruta `WolCsvShare` apunte al inventario real, si se usa.

## Incidencias comunes

### El launcher no actualiza

Revisar:

- acceso a `SharedServerBase`;
- credenciales;
- presencia de `launcher_v*.exe` en la ruta remota.

### La app no encuentra equipos

Revisar:

- que `ComputerNames.sqlite` no este vacia;
- que la importacion CSV haya usado la columna `equipo`;
- que la base local este cerrada antes de sobrescribirla desde fuera.

### Un script no aparece

Revisar:

- `scripts_db.json`;
- que el archivo exista en `scripts\`;
- ejecutar `Actualizar Datos`.

### WOL no localiza la MAC

Revisar:

- el CSV local;
- la ruta UNC configurada para el inventario;
- el formato de nombre del equipo en el CSV.

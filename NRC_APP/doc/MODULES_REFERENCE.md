# Referencia de módulos y secciones

## Módulos

### `DataCollection.psm1`

Responsabilidades:

- lanzar la recogida remota en segundo plano;
- volcar resultados parciales a la interfaz;
- actualizar estados de permisos, sistema operativo, RDP, VNC y WinRM;
- leer la extensión desde `ComputerNames.sqlite`.

### `SharedDataManager.psm1`

Responsabilidades:

- resolver la ruta local y la ruta compartida;
- exponer `appsettings.json`;
- sincronizar `ComputerNames.sqlite`, JSON, scripts e iconos;
- publicar cambios al recurso compartido cuando el usuario lo solicita.

Funciones relevantes:

- `Get-AppSettings`
- `Save-AppSettings`
- `Get-AppSettingValue`
- `Invoke-FullSyncFromServer`
- `Sync-ComputerDbToServer`
- `Sync-ScriptsJsonToServer`
- `Sync-AppsJsonToServer`

### `ScriptRunner.psm1`

Responsabilidades:

- ejecutar scripts según el método configurado;
- preparar copias remotas cuando procede;
- usar `PsExec` en los flujos que requieren `SYSTEM`.

### `DHCP.psm1`

Responsabilidades:

- consultas puntuales sobre DHCP;
- apoyo a la consola DHCP y a tareas administrativas del entorno.

### `DBAccess.psm1`

Responsabilidades:

- acceso auxiliar a SQLite presente en el árbol.

## Secciones

### `FerramentasAdmin.psm1`

Construye el menú `Administrador`. Contiene accesos a:

- consolas del sistema;
- portal principal;
- portal secundario;
- DHCP;
- WSUS;
- AD;
- impresoras;
- AdExplorer.

### `LocalHost.psm1`

Construye el menú `LocalHost`. Incluye:

- información y propiedades del sistema;
- herramientas MMC del equipo local;
- netstat y snap-ins;
- restauración de credenciales VNC;
- otras aplicaciones de Windows.

### `Scripts.psm1`

Construye el menú `Scripts`. Se encarga de:

- registrar scripts integrados;
- leer `scripts_db.json`;
- añadir, modificar o eliminar scripts personalizados;
- replicar scripts e iconos al recurso compartido.

### `Aplicacions.psm1`

Construye el menú `Aplicacions`. Actualmente gestiona:

- `Chat Remoto`;
- `ExtFinder`;
- aplicaciones externas registradas en `apps_db.json`.

### `Configuracion.psm1`

Construye el menú `Configuracion`. Incluye:

- editor de equipos;
- editor de entorno;
- limpieza de Pass Keeper.

### `PassKeeper.psm1`

Responsabilidades:

- panel lateral para guardar valores rápidos cifrados;
- alta, copia y borrado de entradas;
- limpieza completa desde `Configuracion`.

## Subaplicaciones

### `app/Chat/`

Chat entre técnico y usuario basado en un archivo compartido y `PsExec`.

### `app/ExtFinder/`

Interfaz WinForms para consultar y mantener la tabla `extensions`.

### `app/WOL/`

Herramienta VBScript para Wake on LAN con soporte de relay entre subredes.

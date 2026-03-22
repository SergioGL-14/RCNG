# Estructura del proyecto

## Nivel superior

```text
RCNG/
|-- Launcher_RNC/
`-- NRC_APP/
```

## `Launcher_RNC/Launcher_RNC/`

| Ruta | Uso |
|---|---|
| `Launcher_RNCG.ps1` | Instalación, actualización, credenciales y arranque |
| `launcher.ps1` | Lanzador auxiliar |
| `NRC.ico` | Icono del launcher |

## `NRC_APP/`

### Raíz

| Archivo | Uso |
|---|---|
| `LazyWinAdmin_v8.0.ps1` | Aplicación principal |
| `launcher_v8.0.exe` | Ejecutable compilado de la aplicación |
| `Explorer++.exe` | Explorador para shares administrativos |
| `NRC.ico` | Icono principal |
| `README.md` | Descripción general de la carpeta |

### `modules/`

| Archivo | Uso |
|---|---|
| `DataCollection.psm1` | Recogida remota de datos |
| `SharedDataManager.psm1` | Sincronización y configuración |
| `ScriptRunner.psm1` | Ejecución de scripts |
| `DHCP.psm1` | Consultas DHCP |
| `DBAccess.psm1` | Acceso auxiliar a SQLite |

### `sections/`

| Archivo | Uso |
|---|---|
| `FerramentasAdmin.psm1` | Menú Administrador |
| `LocalHost.psm1` | Menú LocalHost |
| `Scripts.psm1` | Menú Scripts |
| `Aplicacions.psm1` | Apps internas y externas |
| `Configuracion.psm1` | Equipos, entorno y Pass Keeper |
| `PassKeeper.psm1` | Panel lateral |

### `app/`

| Carpeta | Uso |
|---|---|
| `Chat/` | Chat remoto |
| `ExtFinder/` | Directorio de extensiones |
| `WOL/` | Wake on LAN |

### `database/`

| Archivo | Uso |
|---|---|
| `ComputerNames.sqlite` | Inventario local de equipos y extensiones |
| `scripts_db.json` | Catálogo de scripts |
| `apps_db.json` | Apps externas |
| `appsettings.json` | Parámetros del entorno |

### `scripts/`

Scripts auxiliares e integrados utilizados por el menú `Scripts` y por distintos flujos de la aplicación.

### `tools/`

Dependencias externas utilizadas por la aplicación, como `PsExec.exe`, `AdExplorer` y herramientas auxiliares de soporte.

### `libs/`

| Archivo | Uso |
|---|---|
| `System.Data.SQLite.dll` | Proveedor SQLite |

### `vnc/`

| Archivo | Uso |
|---|---|
| `vncviewer.exe` | Cliente VNC |
| `options.vnc` | Configuración base del visor |

### `icos/`

Iconos `.ico` y recursos serializados usados por la interfaz.

### `csv/`

Archivos de ejemplo para importación, incluido `equipos_ejemplo.csv`.

### `doc/`

Documentación técnica y manual de usuario.

## Observaciones

- `ComputerNames.sqlite` es la base local de trabajo para equipos y extensiones.
- `appsettings.json` centraliza la configuración de entorno.
- `scripts_db.json` y `apps_db.json` permiten ampliar menús sin tocar la estructura base del código.

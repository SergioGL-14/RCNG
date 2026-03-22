# Arquitectura de RCNG

## Visión general

RCNG se organiza en dos capas principales:

1. `Launcher_RNC`, encargado del despliegue, la actualización y el arranque.
2. `NRC_APP`, que contiene la aplicación principal y sus recursos funcionales.

```text
Launcher_RNC
  -> solicita credenciales
  -> comprueba la ruta compartida
  -> copia o actualiza la instalación local
  -> genera temp.pass
  -> inicia la aplicación local

NRC_APP
  -> LazyWinAdmin_v8.0.ps1
  -> modules/
  -> sections/
  -> app/
  -> database/
  -> scripts/
  -> tools/, libs/, vnc/, icos/
```

## Componentes principales

### Launcher

`Launcher_RNCG.ps1` realiza estas tareas:

- pedir usuario y contraseña;
- validar el acceso al recurso compartido principal;
- copiar la aplicación en `C:\NRC_APP` en un primer despliegue;
- actualizar la copia local cuando existe una versión más reciente;
- preservar la base local durante la actualización;
- generar `temp.pass` para la sesión;
- arrancar la aplicación local con las credenciales indicadas.

### Aplicación principal

`LazyWinAdmin_v8.0.ps1` actúa como orquestador central. Sus responsabilidades incluyen:

- cargar ensamblados y recursos WinForms;
- abrir la base `ComputerNames.sqlite`;
- importar módulos y secciones;
- construir la interfaz;
- registrar eventos de botones y menús;
- coordinar sincronización, recogida de datos y ejecución de scripts.

### Módulos

| Módulo | Responsabilidad |
|---|---|
| `DataCollection.psm1` | Recogida asíncrona de datos remotos |
| `SharedDataManager.psm1` | Configuración de entorno y sincronización |
| `ScriptRunner.psm1` | Ejecución de scripts locales o remotos |
| `DHCP.psm1` | Consultas directas a DHCP |
| `DBAccess.psm1` | Acceso auxiliar a SQLite presente en el árbol |

### Secciones

| Sección | Responsabilidad |
|---|---|
| `FerramentasAdmin.psm1` | Menú Administrador |
| `LocalHost.psm1` | Menú LocalHost |
| `Scripts.psm1` | Menú Scripts y catálogo dinámico |
| `Aplicacions.psm1` | Apps internas y externas |
| `Configuracion.psm1` | Inventario de equipos y configuración de entorno |
| `PassKeeper.psm1` | Panel lateral de claves rápidas |

### Subaplicaciones activas

| Carpeta | Uso |
|---|---|
| `app/Chat/` | Chat remoto técnico-usuario |
| `app/ExtFinder/` | Directorio de extensiones |
| `app/WOL/` | Encendido Wake on LAN |

## Flujos principales

### Despliegue y arranque

1. El usuario ejecuta el launcher.
2. El launcher solicita credenciales y comprueba la ruta compartida.
3. Si no existe copia local, crea la instalación.
4. Si existe, compara y actualiza la versión local.
5. Genera `temp.pass` y abre la aplicación.

### Recogida de datos

1. El usuario introduce un equipo y pulsa `Recolle Datos`.
2. `DataCollection.psm1` lanza un `Start-Job`.
3. El job consulta acceso administrativo, puertos, CIM/DCOM y datos básicos.
4. Un temporizador WinForms consume el resultado parcial disponible.
5. La interfaz actualiza el panel de salida y los indicadores de estado.

### Sincronización

El botón `Actualizar Datos` llama a `Invoke-FullSyncFromServer` para sincronizar:

- `ComputerNames.sqlite`;
- `scripts_db.json`;
- `apps_db.json`;
- scripts, iconos y otros recursos configurados.

Después de la sincronización, la aplicación reconstruye los menús dinámicos y reabre el acceso a la base local.

### Ejecución de scripts

`Scripts.psm1` combina definiciones integradas con el contenido de `scripts_db.json`. Cada entrada se resuelve mediante uno de estos métodos:

- `standard`
- `psexec-system`
- `batch-remote`
- `local`
- `custom`

### Configuración de entorno

`Configuracion -> Entorno global` edita `database/appsettings.json`. Desde ese diálogo se administran:

- la ruta compartida principal;
- la URL PAC del proxy;
- los portales web;
- la ruta y el nombre del CSV usado por WOL;
- el servidor DHCP por defecto;
- el nombre y correo visibles del soporte;
- las bases LDAP usadas por los formularios de grupos.

## Persistencia

| Recurso | Uso |
|---|---|
| `ComputerNames.sqlite` | Equipos y extensiones |
| `scripts_db.json` | Catálogo de scripts |
| `apps_db.json` | Catálogo de apps externas |
| `appsettings.json` | Parámetros del entorno |
| `temp.pass` | Secreto temporal de sesión para VNC |
| `%LOCALAPPDATA%\LazyWinAdmin\pk_*` | Datos cifrados de Pass Keeper |

## Canales técnicos

| Canal | Uso |
|---|---|
| CIM/DCOM y WMI | Recogida de información remota |
| SMB/UNC | Copias, sincronización y chat remoto |
| PsExec | Ejecución remota y elevación a `SYSTEM` |
| RDP y VNC | Acceso remoto |
| RSAT DHCP | Consola y consultas DHCP |

## Consideraciones

- RCNG trabaja con copia local y sincronización posterior.
- La configuración de entorno se gestiona desde `appsettings.json` y desde la interfaz.
- La aplicación principal sigue concentrando buena parte de la interfaz y del orquestado general.

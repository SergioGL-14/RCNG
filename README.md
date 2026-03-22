# RCNG

RCNG es una herramienta de soporte remoto para entornos Windows desarrollada sobre PowerShell 5.1 y WinForms. El proyecto concentra en una sola interfaz tareas habituales de administración, inventario, acceso remoto, ejecución de scripts, mantenimiento de datos locales y utilidades auxiliares de soporte.

El repositorio se organiza en dos bloques principales:

- [`Launcher_RNC/`](Launcher_RNC/), responsable del despliegue, la actualización y el arranque.
- [`RNCG/`](NRC_APP/), que contiene la aplicación principal, sus módulos, secciones, datos y documentación.

## Origen del proyecto

RCNG es un fork de [LazyWinAdmin_GUI](https://github.com/lazywinadmin/LazyWinAdmin_GUI), proyecto publicado por `lazywinadmin` como una utilidad PowerShell con interfaz WinForms orientada a la administración remota de equipos Windows.

La idea original se mantiene: una consola centralizada para operaciones de soporte técnico. A partir de esa base, RCNG amplía el proyecto con un launcher propio, persistencia local, sincronización, configuración de entorno, subaplicaciones integradas y una organización modular del código.

## Qué aporta RCNG respecto al fork original

Las diferencias principales frente a `LazyWinAdmin_GUI` son:

- separación entre launcher y aplicación principal;
- estructura modular mediante `modules/` y `sections/`;
- persistencia local en SQLite y catálogos en JSON;
- configuración de entorno editable desde la interfaz;
- sincronización entre copia local y recurso compartido configurable;
- catálogo dinámico de scripts y aplicaciones externas;
- subaplicaciones integradas como `Chat`, `ExtFinder` y `WOL`;
- almacenamiento cifrado de datos de sesión y `Pass Keeper`;
- adaptación del proyecto a un flujo de trabajo operativo con despliegue local.

RCNG no replica la estructura del repositorio original, sino que toma su base funcional y la desarrolla como una aplicación más amplia y más organizada.

## Funcionalidad principal

Desde la interfaz principal es posible:

- consultar información remota de equipos Windows;
- validar conectividad, permisos administrativos, RDP, VNC y WinRM;
- abrir `RDP`, `CMD Remota`, `PS Remota`, `VNC` y `Explorer++`;
- ejecutar acciones rápidas como `GPUpdate`, cierre de sesión, reinicio, encendido por WOL y apagado;
- acceder a herramientas administrativas desde el menú `Administrador`;
- abrir consolas y herramientas del puesto técnico desde `LocalHost`;
- lanzar scripts integrados o personalizados desde `Scripts`;
- abrir subaplicaciones internas y aplicaciones externas registradas en `Aplicacions`;
- mantener el inventario local de equipos y extensiones;
- editar la configuración global del entorno desde `Configuracion`.

## Arquitectura actual

### `RNCG`

El launcher gestiona el ciclo de arranque:

1. solicita credenciales;
2. comprueba la disponibilidad del recurso compartido configurado;
3. copia o actualiza la instalación local en `C:\RNCG`;
4. genera `temp.pass` para la sesión;
5. lanza la aplicación local.

### `RNCG`

La aplicación principal parte de [`LazyWinAdmin_v8.0.ps1`](RNCG/LazyWinAdmin_v8.0.ps1), que actúa como orquestador de interfaz, datos y acciones remotas. A su alrededor se distribuyen:

- `modules/` para lógica transversal;
- `sections/` para menús y áreas funcionales;
- `database/` para SQLite y JSON;
- `app/` para subaplicaciones;
- `scripts/`, `tools/`, `libs/`, `vnc/` e `icos/` como recursos operativos.

## Componentes destacados

### Módulos

- [`SharedDataManager.psm1`](NRC_APP/modules/SharedDataManager.psm1): configuración de entorno, rutas compartidas y sincronización.
- [`DataCollection.psm1`](NRC_APP/modules/DataCollection.psm1): recogida asíncrona de información remota.
- [`ScriptRunner.psm1`](NRC_APP/modules/ScriptRunner.psm1): ejecución de scripts según el método configurado.
- [`DHCP.psm1`](NRC_APP/modules/DHCP.psm1): consultas DHCP cuando el puesto dispone de RSAT.

### Secciones

- [`FerramentasAdmin.psm1`](NRC_APP/sections/FerramentasAdmin.psm1): acceso a consolas administrativas y portales.
- [`LocalHost.psm1`](NRC_APP/sections/LocalHost.psm1): herramientas del equipo local.
- [`Scripts.psm1`](NRC_APP/sections/Scripts.psm1): catálogo de scripts integrados y personalizados.
- [`Aplicacions.psm1`](NRC_APP/sections/Aplicacions.psm1): subaplicaciones internas y apps externas.
- [`Configuracion.psm1`](NRC_APP/sections/Configuracion.psm1): inventario de equipos, parámetros de entorno y limpieza de Pass Keeper.
- [`PassKeeper.psm1`](NRC_APP/sections/PassKeeper.psm1): almacenamiento cifrado de valores de uso rápido.

### Subaplicaciones

- [`app/Chat/`](NRC_APP/app/Chat/): chat remoto entre técnico y usuario.
- [`app/ExtFinder/`](NRC_APP/app/ExtFinder/): consulta y mantenimiento de extensiones.
- [`app/WOL/`](NRC_APP/app/WOL/): encendido Wake on LAN con soporte de relay entre subredes.

## Configuración y persistencia

RCNG utiliza varios recursos locales:

- [`ComputerNames.sqlite`](NRC_APP/database/ComputerNames.sqlite) para equipos y extensiones;
- [`scripts_db.json`](NRC_APP/database/scripts_db.json) para el catálogo de scripts;
- [`apps_db.json`](NRC_APP/database/apps_db.json) para las aplicaciones externas;
- [`appsettings.json`](NRC_APP/database/appsettings.json) para la configuración del entorno.

La configuración global permite definir, entre otros, estos valores:

- ruta compartida principal;
- URL PAC del proxy;
- portales web;
- inventario WOL;
- servidor DHCP por defecto;
- nombre y correo visibles del soporte;
- bases LDAP empleadas por los formularios de grupos.

## Credenciales y cifrado

RCNG utiliza dos mecanismos principales:

- `temp.pass`, generado por el launcher para reutilizar la contraseña de sesión en el flujo VNC;
- `Pass Keeper`, que guarda datos cifrados por usuario en `%LOCALAPPDATA%`.

El fichero `temp.pass` se cifra con una clave derivada del `MachineGuid` del equipo. `Pass Keeper` mantiene su propio esquema de cifrado local para los valores almacenados por el usuario.

## Requisitos técnicos

Los requisitos base del proyecto son:

- Windows con PowerShell 5.1;
- .NET Framework con soporte WinForms;
- [`System.Data.SQLite.dll`](NRC_APP/libs/System.Data.SQLite.dll);
- acceso a red para SMB/UNC, WMI/CIM/DCOM, RDP, VNC e ICMP según la operación realizada.

Algunas funciones dependen además de herramientas o consolas del puesto técnico, como `dsa.msc`, `printmanagement.msc`, `dhcpmgmt.msc`, `wsus.msc` o utilidades incluidas en `tools/`.

## Arranque y uso inicial

El flujo recomendado es:

1. ejecutar el launcher desde [`Launcher_RNCG/`](Launcher_RNC/);
2. desplegar o actualizar la copia local;
3. revisar [`NRC_APP/database/appsettings.json`](NRC_APP/database/appsettings.json) o `Configuracion -> Entorno global`;
4. cargar el inventario de equipos si se va a usar `Actualizar Datos`, `ExtFinder` o WOL.

## Documentación incluida

La documentación funcional y técnica del proyecto está en [`RNCG/doc/`](NRC_APP/doc/):

- [Índice documental](NRC_APP/doc/README.md)
- [Arquitectura](NRC_APP/doc/ARCHITECTURE.md)
- [Estructura del proyecto](NRC_APP/doc/PROJECT_STRUCTURE.md)
- [Dependencias](NRC_APP/doc/DEPENDENCIES.md)
- [Instalación](NRC_APP/doc/INSTALLATION.md)
- [Mantenimiento](NRC_APP/doc/MAINTENANCE.md)
- [Referencia de módulos](NRC_APP/doc/MODULES_REFERENCE.md)
- [Referencia de scripts](NRC_APP/doc/SCRIPTS_REFERENCE.md)
- [Manual de usuario](NRC_APP/doc/USER_GUIDE.md)

## Estructura del repositorio

```text
RCNG/
|-- Launcher_RNC/
`-- NRC_APP/
```

La carpeta `Launcher_RNCG/` contiene el launcher y sus recursos. La carpeta `RNCG/` contiene la aplicación operativa y toda la documentación asociada.

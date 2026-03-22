# Dependencias

## Runtime base

| Componente | Tipo | Uso |
|---|---|---|
| Windows PowerShell 5.1 | Plataforma | Ejecución principal |
| .NET Framework / WinForms | Plataforma | Interfaz gráfica |
| `System.Data.SQLite.dll` | Librería | Acceso a SQLite |

## Herramientas externas incluidas

| Recurso | Ruta habitual | Uso |
|---|---|---|
| `PsExec.exe` | `tools\PsExec.exe` | Ejecución remota y chat |
| `Explorer++.exe` | raíz de `NRC_APP` | Apertura de `C$` |
| `vncviewer.exe` | `vnc\` | Acceso VNC |
| `AdExplorer.exe` | `tools\AdExplorer\` | Utilidad SysInternals |
| `mc-wol.exe` | `app\WOL\` | Wake on LAN |
| `PsExec.exe` adicional | `app\WOL\` | Relay WOL |

## Consolas y componentes opcionales del puesto

Estas funciones dependen de que el puesto técnico disponga de las herramientas de Windows correspondientes:

- `dsa.msc` para Directorio Activo;
- `printmanagement.msc` para administración de impresoras;
- `dhcpmgmt.msc` para consola DHCP;
- `wsus.msc` para WSUS;
- `powershell_ise.exe`, `msinfo32.exe`, `perfmon.msc`, `diskmgmt.msc` y otras consolas MMC del sistema.

## Requisitos de red

| Canal | Uso |
|---|---|
| SMB/UNC | Sincronización, chat remoto y acceso a `C$` |
| CIM/DCOM / WMI | Recogida de datos |
| RDP | Escritorio remoto |
| VNC | Soporte remoto por visor |
| ICMP | Ping y validaciones previas |

## Configuración de entorno

Los parámetros de entorno se administran en `database/appsettings.json` y desde `Configuracion -> Entorno global`.

Claves disponibles:

- `SharedServerBase`
- `ProxyPacUrl`
- `PortalUrl`
- `MailPortalUrl`
- `WolCsvShare`
- `WolCsvFileName`
- `DhcpServer`
- `SupportDisplayName`
- `SupportEmail`
- `PrimaryGroupSearchBase`
- `SecondaryGroupSearchBase`

## Observaciones

- RSAT DHCP y RSAT AD son opcionales según el uso del puesto.
- Si el recurso compartido no está disponible, la aplicación puede seguir funcionando con la copia local existente.

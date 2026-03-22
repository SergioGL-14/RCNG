# Referencia de scripts

Catálogo actual del menú `Scripts`.

## Scripts integrados

| Nombre visible | Archivo | Método | Tipo |
|---|---|---|---|
| Configurar Proxy | Handler interno | `custom` | Integrado |
| Configurar Edge | `Configurar Edge.ps1` | `standard` | Integrado |
| Last Login | `LastLogin.ps1` | `standard` | Integrado |
| Desinstalar KB | `Desinstalar KB.ps1` | `custom` | Integrado |
| Clean Spooler | `Clean Spooler.bat` | `batch-remote` | Integrado |
| Clean Temp | `Clean Temp.bat` | `batch-remote` | Integrado |
| Reset Scardvr | `Reset Scardvr.bat` | `batch-remote` | Integrado |
| Reconectar Lector | `Reconectar Lector.ps1` | `standard` | Integrado |
| Repair Taskbar | `Repair Taskbar.ps1` | `psexec-system` | Integrado |
| Renombrar Perfil | Handler interno | `custom` | Integrado |

## Métodos de ejecución

| Método | Descripción |
|---|---|
| `standard` | Script PowerShell local que recibe `-ComputerName` |
| `psexec-system` | Copia y ejecución remota como `SYSTEM` |
| `batch-remote` | Ejecución de `.bat` o `.cmd` para el equipo destino |
| `local` | Ejecución en el puesto técnico |
| `custom` | Flujo guiado implementado dentro de `Scripts.psm1` |

## Handlers internos

### `Invoke-Scripts_ConfigurarProxy`

- aplica la configuración de proxy usando la URL PAC definida en `appsettings.json`;
- no lanza un archivo externo;
- utiliza los valores configurados en `Entorno global`.

### `Invoke-Scripts_DesinstalarKB`

- solicita el identificador de la KB;
- ejecuta la desinstalación con confirmación;
- usa el flujo guiado de la propia aplicación.

### `Invoke-Scripts_RenombrarPerfil`

- muestra una lista de perfiles detectados en el equipo remoto;
- ejecuta el flujo auxiliar de renombrado;
- confirma la operación antes de aplicar cambios.

## Scripts auxiliares fuera del menú

| Archivo | Uso |
|---|---|
| `List Printers.ps1` | Consulta de impresoras desde la UI |
| `Renombrar Perfil.bat` | Apoyo del flujo de renombrado |
| `Last Login.bat` | Apoyo auxiliar del flujo de consulta |

## Ciclo de vida de un script personalizado

1. Alta desde `Scripts -> Anadir script...`
2. Copia del archivo a `scripts/`
3. Copia opcional del icono a `icos/`
4. Registro en `scripts_db.json`
5. Replicación opcional al recurso compartido

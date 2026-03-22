# Persistencia y datos de RCNG

Resumen de la persistencia local utilizada por RCNG.

## Tipos de persistencia

El proyecto usa cuatro mecanismos:

1. SQLite para inventarios locales.
2. JSON para configuracion dinamica.
3. Ficheros cifrados para secretos de sesion o de usuario.
4. Recursos UNC como origen o destino de sincronizacion.

## SQLite

### `database/ComputerNames.sqlite`

Es la unica base SQLite operativa del proyecto. Contiene:

| Tabla | Uso |
|---|---|
| `computers` | Inventario local de equipos |
| `extensions` | Relacion equipo-extension |

#### Tabla `computers`

| Campo | Tipo | Uso |
|---|---|---|
| `id` | INTEGER | Clave interna |
| `ou` | TEXT | Grupo, sede u OU de referencia |
| `equipo` | TEXT | Nombre del equipo |
| `orig_line` | TEXT | Valor reconstruido para trazabilidad e importacion |

#### Tabla `extensions`

| Campo | Tipo | Uso |
|---|---|---|
| `equipo` | TEXT | Clave principal |
| `extension` | TEXT | Extension asociada |

Estado distribuido actualmente:

- `computers`: 0 registros
- `extensions`: 0 registros

El proyecto incluye la base vacia para dejar una implantacion limpia. Los datos se cargan despues mediante importacion CSV o alta manual.

## JSON

### `database/scripts_db.json`

Define el catalogo del menu Scripts.

Campos principales:

| Campo | Uso |
|---|---|
| `Name` | Nombre visible |
| `FileName` | Archivo fisico |
| `IconFile` | Icono asociado |
| `Category` | Categoria funcional |
| `ExecutionMethod` | Metodo de ejecucion |
| `CustomHandler` | Handler interno cuando no hay archivo directo |
| `BuiltIn` | Marca de entrada integrada |
| `BuiltInKey` | Identificador estable del built-in |
| `SortOrder` | Orden en menu |
| `AddedOn` | Fecha de alta |

### `database/apps_db.json`

Guarda las aplicaciones externas anadidas por el usuario.

Campos:

| Campo | Uso |
|---|---|
| `Name` | Nombre visible |
| `AppPath` | Ruta al ejecutable o script |
| `IconName` | Icono opcional |
| `AddedOn` | Fecha de alta |

### `database/appsettings.json`

Centraliza la configuracion generica del entorno.

Claves actuales:

| Clave | Uso |
|---|---|
| `SharedServerBase` | Ruta UNC del repositorio compartido |
| `ProxyPacUrl` | URL PAC usada por el script de proxy |
| `PortalUrl` | Portal web principal |
| `MailPortalUrl` | Portal web secundario |
| `WolCsvShare` | Ruta UNC del inventario de red usado por WOL |
| `WolCsvFileName` | Nombre del CSV que usa WOL como inventario |
| `DhcpServer` | Servidor DHCP por defecto para consultas RSAT |
| `SupportDisplayName` | Nombre visible en el launcher y en la app |
| `SupportEmail` | Correo mostrado en la informacion de soporte |
| `PrimaryGroupSearchBase` | Base LDAP principal para busquedas de grupos |
| `SecondaryGroupSearchBase` | Base LDAP secundaria para busquedas de grupos |

## Secretos cifrados

### `temp.pass`

- Lo genera `Launcher_RNCG.ps1`.
- Guarda la contrasena de sesion cifrada con una clave derivada del `MachineGuid`.
- Solo esta pensado para reutilizarse en ese mismo equipo.
- La aplicacion principal lo usa al abrir VNC por primera vez.

### Pass Keeper

Ubicacion:

- `%LOCALAPPDATA%\LazyWinAdmin\pk_<hash>\passkeeper.json`

Contenido:

- lista de entradas con identificador, etiqueta y valor cifrado.

Cada valor se cifra con AES-256-CBC y con una clave derivada del `MachineGuid` mas una sal propia del modulo.

## Recursos compartidos

La sincronizacion usa la ruta definida en `SharedServerBase`. El valor distribuido es un ejemplo:

```text
\\server\share\NRC_APP
```

Los elementos que puede sincronizar son:

- `ComputerNames.sqlite`
- `scripts_db.json`
- `apps_db.json`
- `scripts\`
- `icos\`

## Consideraciones operativas

- `ComputerNames.sqlite` es la fuente unica para autocompletado, inventario de equipos y extensiones.
- `ExtFinder` y la recogida de datos comparten la misma tabla `extensions`.
- La base vacia y los CSV de ejemplo forman parte intencionada del proceso de estandarizacion.

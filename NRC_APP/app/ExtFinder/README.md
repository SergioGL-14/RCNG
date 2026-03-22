# ExtFinder — Directorio de Extensiones Telefónicas

Script PowerShell con interfaz gráfica (WinForms) para consultar y mantener un directorio de extensiones telefónicas asociadas a equipos informáticos.

## Archivo

| Archivo | Descripción |
|---------|-------------|
| `ExtFinder.ps1` | Script principal. GUI de búsqueda y CRUD de extensiones. |

## Funcionalidad

- **Búsqueda en tiempo real**: filtra por nombre de equipo o número de extensión (búsqueda parcial con `LIKE`).
- **Listado completo**: botón "Todos" para mostrar todos los registros sin filtro.
- **Añadir extensión**: diálogo para crear un nuevo par equipo/extensión.
- **Modificar extensión**: edita la extensión de un equipo existente (el nombre del equipo queda bloqueado).
- **Eliminar extensión**: borra el registro seleccionado con confirmación.
- **Contador de registros**: pie de ventana muestra número de resultados o total de registros.

## Requisitos

| Dependencia | Ruta |
|-------------|------|
| `System.Data.SQLite.dll` | `libs\System.Data.SQLite.dll` (relativa al repositorio) |
| Base de datos SQLite | `database\ComputerNames.sqlite` |

La tabla utilizada en la base de datos es `extensions`:

```sql
CREATE TABLE extensions (
    equipo    TEXT PRIMARY KEY,
    extension TEXT NOT NULL
);
```

## Uso

```powershell
.\ExtFinder.ps1
```

La ventana (480×400 px, redimensionable) presenta:

1. **Barra de búsqueda** — escribe equipo o extensión y pulsa **Buscar** (o Enter).
2. **Botón Todos** — limpia el filtro y muestra todos los registros.
3. **ListView** — columnas `Equipo` y `Extension`.
4. **Botones inferiores** — **Añadir**, **Modificar**, **Eliminar** sobre el ítem seleccionado.

## Comportamiento de la base de datos

- Conexión perezosa: se abre al primer acceso y se reutiliza durante la sesión.
- Los nombres de equipo se almacenan siempre en **mayúsculas** (`ToUpper()`).
- Las inserciones usan `INSERT OR REPLACE` (upsert por clave primaria `equipo`).

## Contexto

Herramienta integrada en **NRC_APP**. Permite al personal de soporte localizar rápidamente la extensión de un puesto de trabajo sin consultar fuentes externas.

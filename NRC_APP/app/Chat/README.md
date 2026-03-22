# Chat Remoto

Herramienta de chat de texto entre el equipo tecnico y el usuario del equipo remoto sin instalar software adicional de forma permanente. La sesion se apoya en un archivo compartido por UNC y en un script temporal lanzado en el equipo de destino.

## Archivos

| Archivo | Uso |
|---|---|
| `chat.ps1` | Lado tecnico de la sesion |
| `chat_usu.ps1` | Lado usuario. Se copia al `%TEMP%` remoto |
| `Chat.txt` | Plantilla vacia. La sesion crea su propio `chat.txt` en remoto |

## Flujo de funcionamiento

1. La app principal valida el nombre del equipo y la conectividad.
2. Comprueba si existe una sesion de usuario activa.
3. Copia `chat_usu.ps1` al `%TEMP%` del usuario remoto.
4. Crea `chat.txt` en esa misma ruta.
5. Ejecuta `chat_usu.ps1` mediante `PsExec`.
6. Abre `chat.ps1` en el equipo tecnico apuntando al `chat.txt` remoto.

Durante la sesion, ambos extremos leen y escriben sobre el mismo archivo de texto. La actualizacion es periodica y no requiere puertos adicionales ni WinRM.

## Requisitos

- Acceso UNC al share administrativo `\\EQUIPO\C$`.
- Usuario remoto con sesion interactiva iniciada.
- `PsExec.exe` disponible en `tools\PsExec.exe`.
- Permisos para copiar y ejecutar el script temporal en `%TEMP%`.

## Cierre y limpieza

- Al cerrar desde el lado tecnico se puede solicitar el cierre del lado remoto.
- El flujo crea un fichero de senal para cerrar la sesion de forma limpia.
- Si el cierre normal falla, usa el PID remoto como ultimo recurso.
- Los archivos temporales del chat se eliminan al terminar.

## Limitaciones

- El contenido se guarda en texto plano mientras la sesion esta abierta.
- El modelo depende del acceso a `C$`.
- Esta pensado para una sesion por equipo y por usuario en cada momento.

## Relacion con NRC_APP

Esta herramienta se abre desde `Aplicacions -> Chat Remoto`. Es una pieza fija del proyecto y no depende de configuraciones especificas de una organizacion concreta.

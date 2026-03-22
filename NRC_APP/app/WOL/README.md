# WOL - Wake on LAN multi-subred

Script VBScript para encender equipos remotos mediante Wake on LAN. El flujo intenta enviar el paquete magico desde el propio equipo tecnico si origen y destino comparten subred. Si no es asi, busca un relay en la subred remota y ejecuta el envio desde alli.

## Archivos

| Archivo | Uso |
|---|---|
| `WOL.vbs` | Script principal con la logica de ping, busqueda de MAC, relay y verificacion final |
| `mc-wol.exe` | Herramienta que envia el paquete Wake on LAN |
| `PsExec.exe` | Utilidad usada para ejecutar `mc-wol.exe` en el relay |
| `network_inventory_sample.csv` | CSV de ejemplo con nombre, MAC, IP, mascara y subred |

## Fuente de datos

El script trabaja con un CSV de inventario de red. La distribución incluye:

- existe una copia local de ejemplo llamada `network_inventory_sample.csv`;
- puede configurarse una ruta UNC de referencia en `Configuracion -> Entorno global`;
- el valor por defecto es un ejemplo: `\\server\share\network`.

El formato esperado es:

```text
RESOURCE_GUID,Name,MAC,IP,MASK,Subnet
sample-id,EQUIPO-SAMPLE,00:11:22:33:44:55,192:168:10:25,255:255:255:0,192:168:10:0
```

Los campos IP, mascara y subred usan `:` en lugar de `.` en algunos inventarios antiguos. El script normaliza ese formato cuando es necesario.

## Flujo de ejecucion

1. Hace ping al equipo destino.
2. Si responde, informa de que el equipo ya esta encendido y termina.
3. Si no responde, verifica el CSV local y la copia remota de referencia.
4. Obtiene la MAC del equipo destino.
5. Compara la subred local con la subred remota.
6. Si ambas coinciden, envia el paquete desde el equipo local.
7. Si son distintas, busca un relay disponible en la subred remota.
8. Copia `mc-wol.exe` al relay, ejecuta el envio y elimina el binario temporal.
9. Abre una verificacion final por ping.

## Requisitos

- Acceso de lectura al CSV local.
- Acceso UNC a la ruta remota del inventario, si se usa.
- Acceso administrativo al equipo relay cuando el envio necesita salto entre subredes.
- `mc-wol.exe` y `PsExec.exe` en la misma carpeta que `WOL.vbs`.

## Notas operativas

- El boton `Encender` de la aplicacion principal abre una ventana `cmd.exe` y ejecuta `cscript.exe //NOLOGO WOL.vbs NOMBRE_EQUIPO`.
- El flujo esta orientado a redes Windows con shares administrativos disponibles.
- El proyecto ya no distribuye inventarios ligados a una organizacion concreta. Los valores UNC incluidos son solo ejemplos de implantacion.

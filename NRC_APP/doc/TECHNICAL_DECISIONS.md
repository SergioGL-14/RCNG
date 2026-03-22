# Decisiones tecnicas

## Configuracion desacoplada del entorno

Decision:

- mover rutas UNC, proxy y portales a `database/appsettings.json`.

Motivo:

- evitar cambios de codigo para cada implantacion;
- dejar un paquete reutilizable.

## Inventario local en SQLite

Decision:

- mantener `ComputerNames.sqlite` como base local unica para equipos y extensiones.

Motivo:

- buena velocidad local;
- despliegue simple;
- integracion directa con autocompletado y ExtFinder.

## Sincronizacion local-first

Decision:

- seguir trabajando con copia local y sincronizacion opcional con el servidor.

Motivo:

- resiliencia cuando la ruta compartida no esta disponible;
- menor dependencia del estado de la red.

## JSON para catalogos dinamicos

Decision:

- mantener `scripts_db.json` y `apps_db.json` como catalogos editables.

Motivo:

- facilidad para anadir o retirar entradas;
- serializacion simple;
- buena integracion con los menus dinamicos.

## `temp.pass` ligado al equipo

Decision:

- conservar `temp.pass` cifrado con una clave derivada del `MachineGuid`.

Motivo:

- permite abrir VNC sin pedir la contrasena varias veces;
- evita reutilizacion directa del fichero en otro equipo.

Limite aceptado:

- no es un vault corporativo ni un secreto portable.

## Pass Keeper local por usuario

Decision:

- guardar las entradas cifradas en `%LOCALAPPDATA%`.

Motivo:

- separa los datos del usuario de la carpeta distribuida;
- evita exponer secretos en el recurso compartido.

## WOL basado en CSV

Decision:

- mantener WOL sobre CSV y utilidades externas.

Motivo:

- el flujo actual ya resuelve subred, relay y verificacion final;
- no obliga a introducir una segunda base de datos solo para encendido.

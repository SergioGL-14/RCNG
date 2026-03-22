# Instalación y arranque

## Rutas base

El launcher trabaja con estos valores de referencia:

| Elemento | Valor |
|---|---|
| Instalacion local | `C:\NRC_APP` |
| Repositorio compartido | `\\server\share\NRC_APP` |

Los valores reales pueden ajustarse desde `Configuracion -> Entorno global`.

## Primer arranque

1. Ejecutar `Launcher_RNC`.
2. Introducir usuario y contrasena.
3. El launcher comprueba si existe copia local.
4. Si no existe, copia toda la aplicacion desde el recurso compartido.
5. Crea `temp.pass`.
6. Lanza el ejecutable local.

## Actualizacion

Si ya existe una copia local:

1. El launcher compara la version local y la remota.
2. Si la remota es mas nueva, ejecuta `robocopy`.
3. En la actualizacion preserva `*.sqlite*` para no sobrescribir datos locales.
4. Regenera `temp.pass`.
5. Lanza la nueva version local.

## Primeros ajustes recomendados

Despues del despliegue inicial conviene revisar:

1. `Configuracion -> Entorno global`
2. `Configuracion -> Equipos`
3. `Aplicacions -> ExtFinder`, si se van a mantener extensiones

Valores a personalizar en `Entorno global`:

- ruta UNC del repositorio compartido;
- URL del proxy PAC;
- portal principal;
- portal secundario;
- ruta UNC del CSV de WOL;
- nombre del CSV de WOL;
- servidor DHCP por defecto;
- nombre y correo visibles del soporte;
- bases LDAP para grupos.

## Carga de datos inicial

El proyecto se entrega con:

- `ComputerNames.sqlite` vacía;
- `apps_db.json` sin apps externas;
- `equipos_ejemplo.csv` como plantilla de importacion;
- `network_inventory_sample.csv` como ejemplo para WOL.

Para poblar el inventario de equipos:

1. abrir `Configuracion -> Equipos`;
2. importar un CSV con columna `equipo` y, opcionalmente, `ou`;
3. guardar los cambios;
4. replicar al recurso compartido si procede.

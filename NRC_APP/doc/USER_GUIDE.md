# Manual de usuario de RCNG

Manual funcional de RCNG. El formato está pensado para lectura en Markdown y para exportación posterior a PDF.

## Indice

- [1. Introduccion](#1-introduccion)
- [2. Como iniciar la aplicacion](#2-como-iniciar-la-aplicacion)
- [3. Que hace el launcher](#3-que-hace-el-launcher)
- [4. Interfaz principal](#4-interfaz-principal)
- [5. Campo de equipo y recogida de datos](#5-campo-de-equipo-y-recogida-de-datos)
- [6. Botones principales de la pestana General](#6-botones-principales-de-la-pestana-general)
- [7. Boton Actualizar Datos](#7-boton-actualizar-datos)
- [8. Menu Administrador](#8-menu-administrador)
- [9. Menu LocalHost](#9-menu-localhost)
- [10. Menu Scripts](#10-menu-scripts)
- [11. Menu Aplicacions](#11-menu-aplicacions)
- [12. Menu Configuracion](#12-menu-configuracion)
- [13. Pestana Equipo y Sistema Operativo](#13-pestana-equipo-y-sistema-operativo)
- [14. Pestanas de consulta tecnica](#14-pestanas-de-consulta-tecnica)
- [15. Como actualizar nombres de equipos](#15-como-actualizar-nombres-de-equipos)
- [16. Como actualizar extensiones](#16-como-actualizar-extensiones)
- [17. Como configurar el entorno](#17-como-configurar-el-entorno)
- [18. Como funciona el encendido de equipos](#18-como-funciona-el-encendido-de-equipos)
- [19. Como funciona la encriptacion](#19-como-funciona-la-encriptacion)
- [20. Como anadir scripts y aplicaciones](#20-como-anadir-scripts-y-aplicaciones)
- [21. Referencias utiles](#21-referencias-utiles)
- [22. Incidencias frecuentes](#22-incidencias-frecuentes)

## 1. Introduccion

RCNG es una herramienta de soporte remoto para Windows. Desde una sola ventana permite consultar equipos, abrir accesos remotos, ejecutar scripts, mantener inventarios locales y lanzar herramientas auxiliares.

La aplicacion esta pensada para trabajar con una copia local, pero puede sincronizar datos con un recurso compartido configurable.

## 2. Como iniciar la aplicacion

El arranque recomendado es:

1. abrir `Launcher_RNC`;
2. introducir usuario y contrasena;
3. dejar que el launcher valide o actualice la copia local;
4. esperar a que se abra la app principal.

Tambien es posible ejecutar directamente `LazyWinAdmin_v8.0.ps1` o `launcher_v8.0.exe` para pruebas o mantenimiento.

## 3. Que hace el launcher

El launcher no es un acceso directo simple. Realiza estas tareas:

1. comprueba si existe la instalacion local en `C:\NRC_APP`;
2. valida el acceso al recurso compartido;
3. copia o actualiza la aplicacion;
4. preserva los `.sqlite` locales cuando actualiza;
5. genera `temp.pass`;
6. lanza el ejecutable local con las credenciales indicadas.

`temp.pass` sirve para que la aplicacion pueda reutilizar la contrasena de sesion al abrir VNC.

Sugerencia: incluir una captura del launcher y del dialogo de autenticacion.

## 4. Interfaz principal

La ventana se divide en estas zonas:

1. barra superior con menus;
2. caja de nombre de equipo y estados;
3. pestanas operativas;
4. panel de salida principal;
5. panel de logs;
6. lateral de Pass Keeper.

Sugerencia: incluir una captura completa de la ventana principal.

## 5. Campo de equipo y recogida de datos

En `Nome de Equipo` se escribe el equipo de destino. El campo usa autocompletado basado en `ComputerNames.sqlite`.

Al pulsar `Recolle Datos`:

- la aplicacion valida el acceso administrativo;
- lanza una recogida asincrona en segundo plano;
- consulta SO, uptime, usuario interactivo, red, hardware, disco, WinRM y puertos VNC/RDP;
- consulta la extension del equipo si existe en la tabla `extensions`;
- actualiza el panel de salida y los indicadores de estado.

Los indicadores superiores muestran:

- permisos;
- estado de SO;
- uptime;
- RDP;
- VNC;
- WinRM.

## 6. Botones principales de la pestana General

La pestana `General` reune los accesos rapidos de trabajo diario:

- `Ping`: abre una comprobacion de conectividad.
- `RDP`: abre el escritorio remoto.
- `CMD Remota`: abre una consola remota.
- `PS Remota`: abre una sesion remota de PowerShell.
- `VNC`: abre el visor VNC.
- `Ip Config`: consulta la red del equipo remoto.
- `Explorer++`: abre `\\equipo\C$`.
- `GPupdate`: fuerza una actualizacion de directivas.
- `Cerrar Sesion`: intenta cerrar la sesion del usuario remoto.
- `Reinicio`: reinicia el equipo.
- `Encender`: lanza el flujo WOL.
- `Apagado`: apaga el equipo.

En el bloque `Consola Administracion` tambien aparecen:

- `AdminEquipos`;
- `Servicios`;
- `Visor Eventos`;
- `Registro`.

Sugerencia: incluir una captura de la pestana General.

## 7. Boton Actualizar Datos

`Actualizar Datos` no consulta un equipo. Su funcion es sincronizar el entorno local de la aplicacion.

Al pulsarlo:

- intenta descargar datos desde `SharedServerBase`;
- refresca `ComputerNames.sqlite`;
- vuelve a leer `scripts_db.json` y `apps_db.json`;
- reconstruye los menus `Scripts` y `Aplicacions`;
- deja trazas en el panel de logs.

Conviene usarlo cuando:

- se hayan cambiado scripts o iconos;
- se hayan anadido aplicaciones externas compartidas;
- se hayan actualizado equipos desde otro puesto.

## 8. Menu Administrador

El menu `Administrador` agrupa accesos del puesto tecnico y enlaces del entorno.

Entradas actuales:

- `Administracion Impresoras`
- `Directorio Activo`
- `Exchange`
- `DHCP`
- `WSUS`
- `Portal Web`
- `Consola de Comandos`
- `Consola Powershell`
- `Editor Poweshell`
- `SysInternals -> AdExplorer`

Notas de uso:

- `Portal Web` abre la URL configurada en `PortalUrl`.
- `Exchange` abre la URL configurada en `MailPortalUrl`.
- `DHCP` y `WSUS` dependen de que esas consolas existan en el puesto.

Sugerencia: incluir una captura del menu `Administrador` desplegado.

## 9. Menu LocalHost

`LocalHost` contiene herramientas del propio puesto tecnico.

Entradas principales:

- `Informacion do Sistema`
- `Propiedades do Sistema`
- `Administrador De Equipo`
- `Administrador De Dispositivos`
- `Administrador De Tareas`
- `Servicios`
- `Registro`
- `Microsoft Management Console`
- `Estadisticas de Rede | Portos Escoitando`
- `Snapins - Modulos Disponibles`
- `Reintroducir Credenciais VNC`
- `Outras Aplicacions de Windows`

`Outras Aplicacions de Windows` agrupa:

- `Administracion De Discos`
- `Directivas De Seguridad`
- `Directivas Locales`
- `Monitor de Rendemento`
- `Programador De Tareas`
- `Almacen de Certificados`
- `Usuarios & Grupos Locales`
- `Recursos Compartidos`

## 10. Menu Scripts

El menu `Scripts` mezcla entradas integradas y scripts custom registrados por el usuario.

Scripts integrados actuales:

- `Configurar Proxy`
- `Configurar Edge`
- `Last Login`
- `Desinstalar KB`
- `Clean Spooler`
- `Clean Temp`
- `Reset Scardvr`
- `Reconectar Lector`
- `Repair Taskbar`
- `Renombrar Perfil`

Puede haber ademas scripts custom, por ejemplo `SFC DISM`, si se registran en el equipo.

### Metodos de ejecucion

- `standard`: script local que recibe `-ComputerName`.
- `psexec-system`: copia y ejecucion remota como `SYSTEM`.
- `batch-remote`: ejecucion de `.bat` o `.cmd`.
- `custom`: flujo guiado desde la propia aplicacion.

### Ejemplos relevantes

`Configurar Proxy` toma la URL PAC configurada en `Configuracion -> Entorno global`.

`Renombrar Perfil` abre un flujo guiado para localizar y renombrar perfiles de usuario en remoto.

Sugerencia: incluir una captura del menu `Scripts` y otra del dialogo `Anadir script`.

## 11. Menu Aplicacions

`Aplicacions` mezcla herramientas internas con aplicaciones externas registradas por el usuario.

Entradas internas actuales:

- `Chat Remoto`
- `ExtFinder`

### Chat Remoto

Abre una conversacion entre tecnico y usuario:

1. valida conectividad;
2. detecta la sesion activa del usuario;
3. copia `chat_usu.ps1` al `%TEMP%` remoto;
4. crea `chat.txt`;
5. lanza el lado remoto con `PsExec`;
6. abre el lado tecnico.

### ExtFinder

Abre el directorio de extensiones. Permite:

- buscar por equipo;
- buscar por extension;
- anadir;
- modificar;
- eliminar.

### Aplicaciones externas

El usuario puede anadir accesos a ejecutables o scripts propios. Esas entradas se guardan en `apps_db.json`.

Sugerencia: incluir una captura del menu `Aplicacions`.

## 12. Menu Configuracion

`Configuracion` concentra el mantenimiento local del producto.

Entradas actuales:

- `Equipos`
- `Entorno global`
- `Limpiar Pass Keeper`

### Equipos

Abre el editor de `ComputerNames.sqlite`. Permite:

- buscar;
- paginar;
- anadir;
- editar;
- eliminar;
- importar CSV;
- guardar;
- replicar la base al recurso compartido.

### Entorno global

Permite editar:

- ruta UNC del repositorio compartido;
- URL PAC del proxy;
- portal principal;
- portal secundario;
- share UNC del inventario WOL;
- nombre del CSV usado por WOL;
- servidor DHCP por defecto;
- nombre visible del soporte;
- correo visible del soporte;
- base LDAP principal para grupos;
- base LDAP secundaria para grupos.

### Limpiar Pass Keeper

Elimina todas las entradas guardadas por el usuario en el panel lateral.

Sugerencia: incluir una captura del dialogo `Configuracion -> Entorno global`.

## 13. Pestana Equipo y Sistema Operativo

Esta pestana agrupa cinco bloques:

### Hardware

- `Placa Base`
- `Procesador`
- `Memoria`
- `Sistema`
- `Impresoras`
- `Dispositivos USB`

### S.O / Software

- `Aplicaciones`
- `PageFile`
- `StartUp`
- `Descripcion de Equipo -> Consultar / Modificar`
- `Escritorio Remoto -> Activa / Desactiva`
- `WinRM -> Activar / Deshabilitar`

### Windows Update

- `Ultimos Updates`

### Usuarios y Grupos

- `Usuarios Locales`
- `Grupos Locales`

### Grupos de Dominio

- `Mostrar Grupos`
- `Agregar a Grupo`

## 14. Pestanas de consulta tecnica

### Red

Incluye comprobaciones de conectividad, IPConfig, puertos, rutas, Tracert, NsLookup, Ping y PathPing.

### Procesos

Incluye listados, propietarios, vista grid, procesos recientes, filtro de memoria y cierre de procesos.

### Servicios

Incluye consulta, arranque, parada, reinicio, vistas completas y filtros sobre servicios.

### Disk Drives

Incluye capacidad, particiones, discos logicos, discos fisicos, relaciones y unidades mapeadas.

### Shares

Incluye MMC de recursos compartidos y listados simples o en grid.

### Event Log

Incluye historial de reinicios, visor de eventos y nombres de logs.

### ExternalTools

Incluye accesos como:

- `Mostrar Grupos`
- `Agregar a Grupo`
- `Rwinsta`
- `Qwinsta`
- `MsInfo32`
- `DriverQuery`
- `SystemInfo`
- `PAExec`
- `PsExec`

## 15. Como actualizar nombres de equipos

Procedimiento recomendado:

1. abrir `Configuracion -> Equipos`;
2. importar un CSV o editar manualmente;
3. asegurarse de que la columna `equipo` este informada;
4. guardar;
5. replicar al recurso compartido si el cambio debe compartirse.

El proyecto incluye `csv\equipos_ejemplo.csv` como plantilla.

## 16. Como actualizar extensiones

Las extensiones se mantienen desde `Aplicacions -> ExtFinder`.

Procedimiento:

1. abrir `ExtFinder`;
2. buscar por equipo o extension;
3. anadir, modificar o eliminar;
4. cerrar la ventana al terminar.

La recogida de datos mostrara la extension cuando el equipo exista en la tabla `extensions`.

## 17. Como configurar el entorno

Abrir `Configuracion -> Entorno global` y revisar estos valores:

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

Los valores distribuidos son ejemplos. Deben sustituirse por los del servidor real.

## 18. Como funciona el encendido de equipos

El boton `Encender` abre una ventana `cmd.exe` y ejecuta `cscript.exe //NOLOGO WOL.vbs EQUIPO`.

El flujo:

1. comprueba si el equipo responde;
2. revisa el CSV local y, si existe, la copia remota;
3. obtiene la MAC;
4. decide si el envio puede hacerse desde la subred local;
5. si no, busca un relay en la subred remota;
6. lanza el paquete Wake on LAN;
7. deja una ventana de seguimiento por ping.

El inventario de red puede vivir en una ruta UNC configurada en `WolCsvShare`.

Sugerencia: incluir una captura de la ventana `cmd` del flujo WOL.

## 19. Como funciona la encriptacion

### `temp.pass`

- lo crea el launcher;
- guarda la contrasena de sesion cifrada;
- la clave se deriva del `MachineGuid`;
- se usa para reutilizar la contrasena al abrir VNC.

### Pass Keeper

- guarda valores cortos del usuario;
- cada valor se cifra con AES-256-CBC;
- el fichero se guarda en `%LOCALAPPDATA%`;
- las entradas pueden copiarse rapidamente al portapapeles desde la interfaz.

### Limitacion importante

Ambos mecanismos estan ligados al equipo local. No deben tratarse como un cofre corporativo compartido.

## 20. Como anadir scripts y aplicaciones

### Anadir un script

1. abrir `Scripts -> Anadir script...`;
2. elegir el archivo;
3. indicar nombre visible;
4. escoger metodo de ejecucion;
5. seleccionar icono si se desea;
6. guardar.

El registro queda en `scripts_db.json`.

### Anadir una aplicacion externa

1. abrir `Aplicacions -> Anadir nuevas aplicaciones...`;
2. indicar nombre;
3. indicar ruta del ejecutable o script;
4. indicar icono si se desea;
5. guardar.

El registro queda en `apps_db.json`.

En ambos casos la aplicacion puede preguntar si se desea replicar el cambio al recurso compartido.

## 21. Referencias utiles

Para ampliar informacion conviene consultar tambien:

- [README de Chat Remoto](../app/Chat/README.md)
- [README de ExtFinder](../app/ExtFinder/README.md)
- [README de WOL](../app/WOL/README.md)
- [Referencia del menu Scripts](SCRIPTS_REFERENCE.md)
- [Arquitectura funcional](ARCHITECTURE.md)

## 22. Incidencias frecuentes

### No aparece informacion del equipo

Revisar:

- nombre correcto;
- conectividad;
- acceso a `\\equipo\C$`;
- permisos remotos;
- estado de WMI/CIM.

### VNC pide contrasena

Puede significar:

- que `temp.pass` no existe;
- que no pudo descifrarse;
- que aun no se ha cargado la contrasena de sesion.

Alternativa:

- usar `LocalHost -> Reintroducir Credenciais VNC`.

### Un script no aparece o no arranca

Revisar:

- `scripts_db.json`;
- metodo de ejecucion configurado;
- presencia del archivo en `scripts\`;
- usar `Actualizar Datos`.

### Chat Remoto no abre la ventana del usuario

Revisar:

- que exista una sesion interactiva;
- acceso a `\\equipo\C$`;
- disponibilidad de `PsExec.exe`.

### WOL no encuentra el equipo

Revisar:

- nombre del equipo en el CSV;
- ruta configurada en `WolCsvShare`;
- formato de MAC, IP y subred.

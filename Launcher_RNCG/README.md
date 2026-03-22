# Launcher_RNC

`Launcher_RNC` reúne los archivos necesarios para desplegar, actualizar y arrancar la instalación local de RCNG.

## Contenido

- [`Launcher_RNC/Launcher_RNCG.ps1`](Launcher_RNC/Launcher_RNCG.ps1): launcher principal.
- [`Launcher_RNC/launcher.ps1`](Launcher_RNC/launcher.ps1): lanzador auxiliar.
- [`Launcher_RNC/NRC.ico`](Launcher_RNC/NRC.ico): icono del launcher.

## Función dentro del proyecto

El launcher:

1. solicita credenciales;
2. valida el acceso al recurso compartido configurado;
3. copia o actualiza la instalación local en `C:\NRC_APP`;
4. genera `temp.pass` para la sesión actual;
5. inicia la aplicación local.

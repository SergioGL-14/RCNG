@echo off
cls

REM Validar que se pase el nombre del equipo como argumento
SET c=%1

IF "%c%"=="" (
    echo [ERROR] No se proporcionó el nombre del equipo.
    echo Uso: script.bat NOMBRE_DEL_EQUIPO
    exit /b
)

echo Verificando conexión con el equipo: %c%
PING %c% -n 1 > NUL
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] No se pudo conectar con el equipo: %c%.
    exit /b
)

REM Eliminar clave del registro en el equipo remoto
echo.
echo [INFO] Eliminando clave del registro en el equipo remoto: %c%...
REG DELETE "\\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\S-1-5-21-532234380-917717105-1845911597-426077" /f
IF %ERRORLEVEL% NEQ 0 (
    exit /b
)
echo [INFO] Clave del registro eliminada exitosamente.

REM Finalizar
echo.
echo [INFO] Proceso completado exitosamente en %c%.
exit /b
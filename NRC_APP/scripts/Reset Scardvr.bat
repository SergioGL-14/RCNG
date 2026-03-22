@echo off
cls

REM El nombre del equipo se pasa como argumento y quitamos comillas si las tiene
SET c=%~1

IF "%c%"=="" (
    echo No se proporcionó el nombre del equipo.
    exit /b
)

REM Verificar conectividad al equipo remoto
PING %c% -n 1 > NUL

IF %ERRORLEVEL%==0 (GOTO CONFIGURAR) ELSE (GOTO ERPING)

:CONFIGURAR
echo.
echo Configurando el servicio de tarjeta inteligente en el equipo remoto '%c%'...
echo.

REM Deshabilitar el servicio en el equipo remoto
sc \\%c% config scardsvr start= disabled

REM Detener el servicio en el equipo remoto
sc \\%c% stop scardsvr

REM Esperar 3 segundos
timeout /t 3 /nobreak >nul

REM Configurar el servicio para inicio automático
sc \\%c% config scardsvr start= auto

REM Iniciar el servicio en el equipo remoto
sc \\%c% start scardsvr

echo.
echo Servicio de tarjeta inteligente reiniciado correctamente en '%c%'.
exit /b

:ERPING
echo Error: No se pudo alcanzar el equipo remoto '%c%'.
exit /b
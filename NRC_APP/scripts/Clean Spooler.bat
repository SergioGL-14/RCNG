@echo off
cls

REM El nombre del equipo se pasa como argumento y quitamos comillas si las tiene
SET c=%~1

IF "%c%"=="" (
    echo No se proporcionó el nombre del equipo.
    exit /b
)

REM El resto de tu código sigue igual, ya que ahora usas %c% como el nombre del equipo
PING %c% -n 1 > NUL

IF %ERRORLEVEL%==0 (GOTO PARAR) ELSE (GOTO ERPING)

:PARAR
echo.
echo Parando servicios LPD, Papercut e Cola de impresion...
echo.
SC \\%c% STOP SPOOLER && ping -n 1 127.0.0.1 >nul

:ESPERA1
PING -n 2 127.0.0.1 >nul
sc \\%c% query SPOOLER | FIND /I "RUNNING"
IF %ERRORLEVEL% == 0 (GOTO PARAR)

echo.
echo Eliminando trabajos en cola...
echo.
del \\%C%\admin$\System32\spool\PRINTERS\*.* /Q/F
echo.

:INICIAR
echo.
echo Iniciando servicios Papercut e Cola de impresion...
echo.
SC \\%c% START SPOOLER && ping -n 1 127.0.0.1 >nul

:ESPERA2
PING -n 2 127.0.0.1 >nul
sc \\%c% query SPOOLER | FIND /I "STOPPED"
IF %ERRORLEVEL% == 0 (GOTO INICIAR)

echo.
exit /b
@echo off
setlocal enabledelayedexpansion

REM Verificar si se pasó un nombre de equipo como argumento
if "%~1"=="" (
    echo No se especificó un nombre de equipo. Por favor, proporcione uno.
    pause
    exit /b
)

REM Variable para verificar si se encontró un usuario
set encontrado=0

REM Ejecutar query user y procesar la salida
for /f "skip=1 tokens=1,2,3,4,* delims= " %%A in ('query user /server:%1 2^>nul') do (
    set usuario=%%A
    set tiempo_inicial=%%E

    REM Filtrar "ninguno" de la salida
    for /f "tokens=*" %%F in ("!tiempo_inicial!") do (
        set tiempo_limpio=%%F
        set tiempo_limpio=!tiempo_limpio:ninguno=!
    )

    REM Si se encontró un tiempo válido, guardar y detener el bucle
    if not "!tiempo_limpio!"=="" (
        set logonTime=!tiempo_limpio!
        set encontrado=1
        goto mostrar_resultado
    )
)

REM Si no se encontró ningún usuario con fecha y hora de inicio de sesión
if "%encontrado%"=="0" (
    echo No se encontró ningún usuario en el equipo remoto: %1
    pause
    exit /b
)

:mostrar_resultado
REM Mostrar resultados finales
echo Usuario: %usuario%
echo Hora de inicio de sesion: %logonTime%
pause
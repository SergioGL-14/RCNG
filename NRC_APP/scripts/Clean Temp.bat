@echo off
cls

REM ============================================================
REM  Limpia temporales y cachés (Edge / Chrome) + flush DNS remoto
REM  Uso:  CleanCacheRemote.bat  EQUIPO
REM  Borra todo por defecto, incluyendo contraseñas guardadas
REM ============================================================

REM --- Nombre del equipo remoto (primer argumento) -------------
SET c=%~1
IF "%c%"=="" (
    echo No se proporcionó el nombre del equipo.
    exit /b
)

REM --- Por defecto, borrar contraseñas (1 = borrar, 0 = no borrar)
SET wipePasswords=1

REM --- Comprobar conectividad ----------------------------------
PING %c% -n 1 >NUL
IF %ERRORLEVEL% NEQ 0 (
    echo Error: No se puede conectar con el equipo remoto %c%.
    exit /b
)

REM --- Identificar usuario con sesión activa -------------------
FOR /F "tokens=1 delims= " %%A IN ('
    query user /SERVER:%c% ^| find /I "activo"
') DO (
    SET loggedInUser=%%A
)

IF "%loggedInUser%"=="" (
    echo No se encontró un usuario activo en el equipo remoto.
    exit /b
)

echo Usuario identificado: %loggedInUser%

REM --- Cerrar navegadores --------------------------------------
echo Cerrando navegadores en el equipo remoto...
FOR %%P IN (chrome.exe msedge.exe firefox.exe brave.exe) DO (
    taskkill /S %c% /IM %%P /F >NUL 2>&1
)

REM --- Limpiar temporales y carpetas clásicas ------------------
echo Limpiando temporales para el usuario %loggedInUser%...
del /S /F /Q \\%c%\C$\Users\%loggedInUser%\AppData\Local\Temp\*.*                        >NUL 2>&1
del /S /F /Q \\%c%\C$\Users\%loggedInUser%\AppData\Local\Microsoft\Windows\INetCache\*.* >NUL 2>&1
rd  /S /Q   \\%c%\C$\Users\%loggedInUser%\AppData\Local\Temp                            >NUL 2>&1

REM --- Limpieza Edge (perfil Default) --------------------------
echo Limpiando cache Edge (perfil Default)...
del /S /F /Q "\\%c%\C$\Users\%loggedInUser%\AppData\Local\Microsoft\Edge\User Data\Default\Cache\*.*"      >NUL 2>&1
del /S /F /Q "\\%c%\C$\Users\%loggedInUser%\AppData\Local\Microsoft\Edge\User Data\Default\GPUCache\*.*"   >NUL 2>&1
del /S /F /Q "\\%c%\C$\Users\%loggedInUser%\AppData\Local\Microsoft\Edge\User Data\Default\Code Cache\*.*" >NUL 2>&1

REM --- Limpieza Chrome (todas las profiles) --------------------
echo Limpiando cache Chrome (todas las profiles)...
FOR /D %%D IN ("\\%c%\C$\Users\%loggedInUser%\AppData\Local\Google\Chrome\User Data\*") DO (
    del /S /F /Q "%%D\Cache\*.*"      >NUL 2>&1
    del /S /F /Q "%%D\GPUCache\*.*"   >NUL 2>&1
    del /S /F /Q "%%D\Code Cache\*.*" >NUL 2>&1
)

REM --- Limpieza Edge (todas las profiles) ----------------------
echo Limpiando cache Edge (todas las profiles)...
FOR /D %%E IN ("\\%c%\C$\Users\%loggedInUser%\AppData\Local\Microsoft\Edge\User Data\*") DO (
    del /S /F /Q "%%E\Cache\*.*"      >NUL 2>&1
    del /S /F /Q "%%E\GPUCache\*.*"   >NUL 2>&1
    del /S /F /Q "%%E\Code Cache\*.*" >NUL 2>&1
)

REM --- Limpieza adicional Edge (History, Cookies, Web Data) ----
echo Limpiando History, Cookies y Web Data de Edge...
FOR /D %%F IN ("\\%c%\C$\Users\%loggedInUser%\AppData\Local\Microsoft\Edge\User Data\*") DO (
    IF EXIST "%%F\History" del /F /Q "%%F\History" >NUL 2>&1
    IF EXIST "%%F\History-journal" del /F /Q "%%F\History-journal" >NUL 2>&1
    IF EXIST "%%F\Cookies" del /F /Q "%%F\Cookies" >NUL 2>&1
    IF EXIST "%%F\Cookies-journal" del /F /Q "%%F\Cookies-journal" >NUL 2>&1
    IF EXIST "%%F\Web Data" del /F /Q "%%F\Web Data" >NUL 2>&1
    IF EXIST "%%F\Web Data-journal" del /F /Q "%%F\Web Data-journal" >NUL 2>&1
    IF EXIST "%%F\Network\Cookies" del /F /Q "%%F\Network\Cookies" >NUL 2>&1
    IF "%wipePasswords%"=="1" (
        IF EXIST "%%F\Login Data" del /F /Q "%%F\Login Data" >NUL 2>&1
        IF EXIST "%%F\Login Data-journal" del /F /Q "%%F\Login Data-journal" >NUL 2>&1
    )
)

REM --- Limpieza adicional Chrome (History, Cookies, Web Data) ---
echo Limpiando History, Cookies y Web Data de Chrome...
FOR /D %%G IN ("\\%c%\C$\Users\%loggedInUser%\AppData\Local\Google\Chrome\User Data\*") DO (
    IF EXIST "%%G\History" del /F /Q "%%G\History" >NUL 2>&1
    IF EXIST "%%G\History-journal" del /F /Q "%%G\History-journal" >NUL 2>&1
    IF EXIST "%%G\Cookies" del /F /Q "%%G\Cookies" >NUL 2>&1
    IF EXIST "%%G\Cookies-journal" del /F /Q "%%G\Cookies-journal" >NUL 2>&1
    IF EXIST "%%G\Web Data" del /F /Q "%%G\Web Data" >NUL 2>&1
    IF EXIST "%%G\Web Data-journal" del /F /Q "%%G\Web Data-journal" >NUL 2>&1
    IF EXIST "%%G\Network\Cookies" del /F /Q "%%G\Network\Cookies" >NUL 2>&1
    IF "%wipePasswords%"=="1" (
        IF EXIST "%%G\Login Data" del /F /Q "%%G\Login Data" >NUL 2>&1
        IF EXIST "%%G\Login Data-journal" del /F /Q "%%G\Login Data-journal" >NUL 2>&1
    )
)

REM --- Flush DNS remoto (sin wmic.exe) -------------------------
echo Haciendo flush de la caché DNS...
powershell -Command "(Get-WmiObject -Class Win32_Process -ComputerName %c%).Create('cmd.exe /c ipconfig /flushdns')" >NUL 2>&1

echo.
echo Limpieza completada en el equipo remoto %c%.
echo Nota: las contrasenas guardadas tambien fueron eliminadas.
exit
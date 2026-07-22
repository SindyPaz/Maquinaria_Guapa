@echo off
chcp 65001 >nul
echo Creando accesos directos en el Escritorio...

set APP1="C:\Users\SindyUrrutia\OneDrive - Agricola Guapa SAS\ANALISTA_OM\09_MAQUINARIA\APPS_SETUP\APP1_Anomalias.html"
set APP2="C:\Users\SindyUrrutia\OneDrive - Agricola Guapa SAS\ANALISTA_OM\09_MAQUINARIA\APPS_SETUP\APP2_Labores.html"
set ESCRITORIO=%USERPROFILE%\Desktop

:: Buscar Chrome
set CHROME="C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist %CHROME% set CHROME="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist %CHROME% (
    echo Chrome no encontrado en rutas default.
    echo Buscando...
    for /f "tokens=*" %%i in ('where chrome 2^>nul') do set CHROME="%%i"
)

:: Crear acceso directo APP1 via PowerShell
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%ESCRITORIO%\Anomalias Maquinaria.lnk'); $sc.TargetPath = %CHROME%; $sc.Arguments = '--app=file:///%APP1:~1,-1%'; $sc.IconLocation = %CHROME% + ',0'; $sc.Description = 'APP1 Anomalias Maquinaria - Agricola Guapa SAS'; $sc.Save()"

:: Crear acceso directo APP2 via PowerShell
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%ESCRITORIO%\Labores Diarias.lnk'); $sc.TargetPath = %CHROME%; $sc.Arguments = '--app=file:///%APP2:~1,-1%'; $sc.IconLocation = %CHROME% + ',0'; $sc.Description = 'APP2 Labores Diarias - Agricola Guapa SAS'; $sc.Save()"

echo.
echo Listo! Revisa tu Escritorio:
echo   - "Anomalias Maquinaria"
echo   - "Labores Diarias"
echo.
echo Doble clic para abrir como app (sin barra del navegador)
pause

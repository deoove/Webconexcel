@echo off
echo Iniciando la aplicacion...

:: Verificar que Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Descargalo de https://www.python.org/downloads/
    echo Asegurate de tildar "Add Python to PATH" al instalar.
    pause
    exit /b
)

:: Instalar dependencias si faltan
echo Verificando dependencias...
pip install flask openpyxl -q --disable-pip-version-check

:: Arrancar el servidor en segundo plano
start "" /b python "%~dp0app.py" > "%~dp0app.log" 2>&1

:: Esperar a que el servidor este listo
echo Abriendo el navegador...
timeout /t 3 /nobreak >nul

:: Abrir el navegador
start http://127.0.0.1:5000

echo.
echo La aplicacion esta corriendo en http://127.0.0.1:5000
echo Cerrá esta ventana para APAGAR la aplicacion.
echo.
pause

:: Al cerrar la ventana, matar el servidor
taskkill /f /im python.exe >nul 2>&1

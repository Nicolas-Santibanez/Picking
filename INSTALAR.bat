@echo off
TITLE Instalando dependencias de Picking...
COLOR 0A

echo ========================================================
echo   Configurando entorno para Generador de Picking...
echo ========================================================

cd /d "%~dp0"

:: Crear el entorno virtual si no existe
if not exist "venv" (
    echo Creando entorno virtual Python...
    py -m venv venv
)

:: Activar e instalar librerías
call venv\Scripts\activate.bat
pip install -r requirements.txt

echo.
echo ========================================================
echo   ¡Instalación completada con éxito!
echo ========================================================
pause
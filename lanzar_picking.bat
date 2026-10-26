@echo off
cd /d "%~dp0"

:: Activar el entorno virtual
call venv\Scripts\activate.bat

:: Iniciar Streamlit y abrir el navegador usando el flag auto-open de Streamlit
streamlit run app.py
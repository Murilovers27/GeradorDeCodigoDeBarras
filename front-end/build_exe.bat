@echo off
setlocal
cd /d "%~dp0.."

if not exist ".venv\Scripts\pyinstaller.exe" (
    echo Instalando PyInstaller no ambiente virtual...
    .venv\Scripts\python.exe -m pip install pyinstaller
)

.venv\Scripts\pyinstaller.exe ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name GeradorEtiquetasABB ^
    --add-data "modelo_etiqueta.docx;." ^
    --paths "." ^
    "front-end\app.py"

echo.
echo Executavel criado em: dist\GeradorEtiquetasABB.exe
pause

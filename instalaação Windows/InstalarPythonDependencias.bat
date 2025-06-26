@echo off
setlocal

echo ===== Verificando Python =====
where python >nul 2>nul
if errorlevel 1 (
    echo Python não encontrado. Instalando...
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
) else (
    echo Python já instalado.
)

echo ===== Instalando dependências =====
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas selenium openpyxl webdriver-manager

echo ===== Gerando executável =====
pyinstaller --noconfirm --onefile --add-data "chamados.xlsx;." executarChamados.py

echo ===== Concluído =====
pause
endlocal

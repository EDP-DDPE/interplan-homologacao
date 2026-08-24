@echo off
setlocal

REM Vai para a pasta onde o bat está
cd /d "%~dp0"

REM Verifica se o requirements existe
if not exist "requirements.txt" (
  echo ERRO: requirements.txt nao encontrado.
  pause
  exit /b 1
)

REM Cria ambiente virtual se nao existir
if not exist ".venv\Scripts\python.exe" (
  echo Criando ambiente virtual...
  py -m venv .venv || python -m venv .venv
)

REM Instala dependencias
echo Instalando dependencias...
".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet --disable-pip-version-check

if errorlevel 1 (
  echo ERRO ao instalar dependencias.
  pause
  exit /b 2
)

REM Executa a homologacao
echo Executando homologacao...
".venv\Scripts\python.exe" main.py

pause

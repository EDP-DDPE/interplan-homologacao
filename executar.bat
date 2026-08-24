@echo off

set REPO=C:\Ferramentas\Interplan-Homologacao

if not exist "%REPO%" (
  echo Repositorio nao encontrado.
  pause
  exit
)

cd /d "%REPO%"

git pull

pip install -r requirements.txt

python main.py "%CD%"

pause

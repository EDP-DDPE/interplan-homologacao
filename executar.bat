@echo off

echo Instalando dependencias...
pip install -r requirements.txt

echo Executando comparacao...
python main.py

echo.
echo Processo concluido.
pause

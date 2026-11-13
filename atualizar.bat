@echo off
cd /d "%~dp0"
python consolidar_presenca.py
if exist saida\painel.html start "" saida\painel.html
pause

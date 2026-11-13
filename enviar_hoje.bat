@echo off
cd /d "%~dp0"
echo Vai ENVIAR os e-mails de presenca aos supervisores e o geral.
choice /c SN /m "Enviar agora"
if errorlevel 2 exit /b
python enviar_emails.py --enviar
pause

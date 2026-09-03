@echo off
REM Executa a automacao de comparacao de dados usando o Python do .venv
cd /d "%~dp0"
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py" >> "%~dp0agendador_execucoes.log" 2>&1

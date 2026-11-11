@echo off
REM Executa a automacao de comparacao de dados usando o Python do .venv
REM O main.py ja registra tudo em automacao_servicos.log; aqui deixamos o
REM progresso aparecer na propria tela tambem (sem redirecionar a saida).
cd /d "%~dp0"
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py"

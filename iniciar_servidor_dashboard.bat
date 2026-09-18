@echo off
REM Sobe o servidor do dashboard em segundo plano (pythonw = sem janela de console).
REM Essa janela do .bat fecha sozinha; o servidor continua rodando "invisível".
REM Pra conferir se subiu certo ou ver os links de acesso: servidor_dashboard.log
REM Pra parar o servidor: Gerenciador de Tarefas > localizar "pythonw.exe" > Finalizar tarefa.
cd /d "%~dp0"
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0servidor_dashboard.py"

@echo off
cd /d "%~dp0"
python atualizar_planilha.py >> "%~dp0log_agendador.txt" 2>&1

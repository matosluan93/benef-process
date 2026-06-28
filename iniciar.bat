@echo off
title BenefProcess - Stefanini

cd /d "%~dp0"

echo.
echo  ============================================
echo   BenefProcess - Tratamento de Beneficios
echo   Stefanini - Gente e Cultura
echo  ============================================
echo.

:: Verifica se o ambiente virtual ja existe
if not exist "venv\Scripts\python.exe" (
    echo  [1/3] Criando ambiente virtual...
    py -m venv venv
    if errorlevel 1 (
        echo.
        echo  ERRO: Python nao encontrado. Instale o Python 3.10+ em python.org
        pause
        exit /b 1
    )
    echo.

    echo  [2/3] Instalando dependencias...
    venv\Scripts\pip install -r requirements.txt --quiet
    echo.
) else (
    echo  Ambiente virtual encontrado. Pulando instalacao.
    echo.
)

echo  [3/3] Iniciando servidor...
echo.
echo  Acesse no navegador: http://localhost:5050
echo  Pressione Ctrl+C para encerrar.
echo.

venv\Scripts\python app.py

pause

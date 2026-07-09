#!/bin/bash
cd "$(dirname "$0")"

echo ""
echo "============================================"
echo "  BenefProcess - Tratamento de Beneficios"
echo "  Stefanini - Gente e Cultura"
echo "============================================"
echo ""

if [ ! -f "venv/Scripts/python.exe" ] && [ ! -f "venv/bin/python" ]; then
    echo "[1/3] Criando ambiente virtual..."
    py -m venv venv 2>/dev/null || python3 -m venv venv
    echo ""
else
    echo "[1/3] Ambiente virtual encontrado."
    echo ""
fi

source venv/Scripts/activate 2>/dev/null || source venv/bin/activate
echo "[2/3] Conferindo dependencias..."
pip install -r requirements.txt --quiet
if [ $? -ne 0 ]; then
    echo ""
    echo "ERRO: Falha ao instalar dependencias."
    exit 1
fi
echo ""

echo "[3/3] Iniciando servidor..."
echo ""
echo "  Acesse: http://localhost:5050"
echo "  Ctrl+C para encerrar."
echo ""

py app.py 2>/dev/null || python app.py

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

    echo "[2/3] Instalando dependencias..."
    source venv/Scripts/activate 2>/dev/null || source venv/bin/activate
    pip install -r requirements.txt --quiet
    echo ""
else
    echo "Ambiente virtual encontrado."
    source venv/Scripts/activate 2>/dev/null || source venv/bin/activate
    echo ""
fi

echo "[3/3] Iniciando servidor..."
echo ""
echo "  Acesse: http://localhost:5050"
echo "  Ctrl+C para encerrar."
echo ""

py app.py 2>/dev/null || python app.py

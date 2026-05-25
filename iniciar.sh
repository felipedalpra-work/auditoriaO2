#!/bin/bash
cd "$(dirname "$0")"
echo "📦 Instalando dependências..."
pip3 install -r requirements.txt -q
echo "🚀 Iniciando servidor O2 Auditoria..."
echo "✅ Acesse: http://localhost:5000"
python3 app.py

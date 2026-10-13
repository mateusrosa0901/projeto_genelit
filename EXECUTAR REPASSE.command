#!/bin/bash
# Mac: dê dois cliques neste arquivo (na primeira vez: botão direito > Abrir).
cd "$(dirname "$0")"
echo "Instalando/verificando bibliotecas (só demora na primeira vez)..."
python3 -m pip install -q -r requirements.txt 2>/dev/null || python3 -m pip install -q --user -r requirements.txt
python3 main.py
echo
read -n 1 -s -r -p "Pressione qualquer tecla para fechar..."

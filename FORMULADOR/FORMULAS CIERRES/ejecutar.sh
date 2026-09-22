#!/bin/bash
cd "$(dirname "$0")"

VENV="../../venv"

if [ ! -d "$VENV" ]; then
    echo "Creando entorno virtual..."
    python3 -m venv "$VENV"
fi

source "$VENV/bin/activate"

if ! python3 -c "import PySide6" 2>/dev/null; then
    echo "Instalando PySide6..."
    pip install PySide6
fi

python3 formula_builder.py
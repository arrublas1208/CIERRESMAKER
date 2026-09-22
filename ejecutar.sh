#!/bin/bash
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
    echo "Creando entorno virtual..."
    python3 -m venv venv
fi

source venv/bin/activate

if ! python3 -c "import PySide6" 2>/dev/null; then
    echo "Instalando PySide6..."
    pip install PySide6
fi

python3 main.py
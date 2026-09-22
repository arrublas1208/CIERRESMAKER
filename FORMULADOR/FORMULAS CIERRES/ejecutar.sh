#!/bin/bash
cd "$(dirname "$0")"

detect_python() {
    for c in python py python3; do
        if command -v "$c" >/dev/null 2>&1; then
            if "$c" -c "import PySide6" >/dev/null 2>&1; then
                echo "$c"
                return 0
            fi
        fi
    done
    # Probar el lanzador de Windows con distintas versiones
    for v in 3.12 3.13 3.11 3.10; do
        if command -v py >/dev/null 2>&1 && py -$v -c "import PySide6" >/dev/null 2>&1; then
            echo "py -$v"
            return 0
        fi
    done
    return 1
}

PY=$(detect_python)
if [ -z "$PY" ]; then
    echo "No se encontro un Python con PySide6 instalado."
    echo "Instalalo con:  python -m pip install PySide6"
    exit 1
fi

echo "Usando: $PY"
$PY formula_builder.py
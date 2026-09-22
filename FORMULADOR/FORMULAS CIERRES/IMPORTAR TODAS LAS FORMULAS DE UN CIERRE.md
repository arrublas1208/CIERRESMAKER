

PROMPT
****
la idea de esto es con la deteccion de formulas con su nombre y sus celdas pues basicamente es importar y asi sin modificar su posicion ni la estrucutura de json que viene si no solo detectar sus formulas poder reconocerlas visuamente con el nombre que si uno presiona una formula el resalta con color las formulas que tocan esta o si se modifica una formula que otra tiene pues el indicara que esta formula se tiene que acrualizar para tener la integridad de las formulas
el archivo de ejemplo es ejemplo.json

1. Entender la estructura del JSON
El JSON tiene esta forma (simplificada):

json
[
  {
    "name": "sheet1",
    "freeze": "A1",
    "styles": [ ... ],
    "merges": ["B5:D5", "B6:D6", ...],
    "rows": {
      "0": { "cells": { "2": { "text": "ACEITES MORICHAL S.A.S", "style": 372 }, ... } },
      "1": { "cells": { "1": { "text": "INFORME DIARIO...", "style": 341, "merge": [0,6] }, ... } },
      ...
    },
    "cols": { "0": { "width": 104 }, ... }
  }
]
rows es un diccionario donde la clave es el índice de fila (0-based).

Cada fila tiene cells, otro diccionario donde la clave es el índice de columna (0-based).

Cada celda puede tener:

"text": el contenido (fórmula o texto).

"style": índice al array de estilos.

"merge": [rowspan, colspan] si es parte de una combinación.

Las fórmulas están en "text" y suelen empezar con fltFLoatTablaDinamica(...) o contener expresiones complejas.

2. Script en Python para extraer fórmulas y posiciones
python
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# 1. Cargar el JSON (si está en un archivo .txt con el string escapado)
with open('datos.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Si el archivo contiene el string escapado, primero hay que desescaparlo:
# contenido = f.read()
# data = json.loads(contenido)

hoja = data[0]  # Primera hoja
rows = hoja.get('rows', {})

# 2. Recorrer todas las celdas y extraer información
celdas = []
for r_idx_str, fila in rows.items():
    r_idx = int(r_idx_str)  # 0-based
    for c_idx_str, celda in fila.get('cells', {}).items():
        c_idx = int(c_idx_str)
        texto = celda.get('text', '')
        estilo = celda.get('style')
        merge = celda.get('merge')  # [rowspan, colspan]
        
        # Convertir a notación A1
        col_letra = get_column_letter(c_idx + 1)
        celda_a1 = f"{col_letra}{r_idx + 1}"
        
        # Detectar si es fórmula (ajusta según tus patrones)
        es_formula = (
            texto.startswith('=') or 
            'fltFLoatTablaDinamica' in texto or
            texto.strip().startswith('$')  # muchas fórmulas empiezan con $
        )
        
        celdas.append({
            'fila': r_idx + 1,
            'columna': col_letra,
            'celda': celda_a1,
            'texto': texto,
            'es_formula': es_formula,
            'estilo': estilo,
            'merge': merge
        })

# 3. Filtrar solo las fórmulas
formulas = [c for c in celdas if c['es_formula']]

print(f"Total de celdas: {len(celdas)}")
print(f"Total de fórmulas: {len(formulas)}")
print("\nPrimeras 10 fórmulas:")
for f in formulas[:10]:
    print(f"{f['celda']}: {f['texto'][:120]}...")

# 4. Exportar todas las celdas a Excel respetando posiciones
wb = Workbook()
ws = wb.active
ws.title = "Extraido"

for c in celdas:
    ws[f"{c['columna']}{c['fila']}"] = c['texto']

wb.save('hoja_extraida.xlsx')
print("\nArchivo Excel generado: hoja_extraida.xlsx")

# 5. Opcional: guardar solo fórmulas en un JSON más limpio
with open('formulas.json', 'w', encoding='utf-8') as f:
    json.dump(formulas, f, ensure_ascii=False, indent=2)
print("Fórmulas guardadas en formulas.json")
3. ¿Qué hace este script?
Carga el JSON (si está escapado, se puede desescapar con json.loads sobre el string).

Recorre fila por fila y celda por celda usando los índices numéricos.

Convierte los índices a notación A1 (ej. fila 5, columna 2 → C6).

Detecta fórmulas según patrones comunes (fltFLoatTablaDinamica, $, =). Puedes ajustar la condición.

Exporta a Excel con openpyxl, colocando cada texto en su celda exacta.

Guarda un JSON solo con las fórmulas y sus posiciones para análisis posterior.

4. Consideraciones adicionales
Celdas combinadas (merges): En hoja['merges'] tienes rangos como "B5:D5". Solo la celda superior izquierda contiene el texto; las demás están vacías. Si necesitas expandir el valor a todas las celdas del merge, puedes parsear esos rangos y copiar el texto.

Estilos: El campo style es un índice al array styles. Si necesitas conservar formatos, tendrías que aplicar el estilo correspondiente en Excel, pero para extraer fórmulas no es imprescindible.

Fórmulas de x-spreadsheet: No son fórmulas nativas de Excel, por lo que al exportar a .xlsx se guardarán como texto. Si quieres que Excel las evalúe, tendrías que traducirlas, lo cual es complejo. Para extraerlas y analizarlas, guardarlas como texto es lo correcto.

Archivo de entrada: Si el contenido que mostraste está dentro de un .txt con las comillas escapadas, primero haz:

python
with open('archivo.txt', 'r', encoding='utf-8') as f:
    contenido = f.read()
data = json.loads(contenido)
5. Resultado
Obtendrás:

Un archivo Excel (hoja_extraida.xlsx) con todas las celdas en su posición original.

Un archivo JSON (formulas.json) con cada fórmula, su celda A1, fila, columna, texto completo, etc.

Así podrás ubicar exactamente dónde está cada fórmula y trabajar con ellas desde Python.




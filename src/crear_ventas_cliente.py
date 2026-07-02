from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"

DATA_DIR.mkdir(exist_ok=True)

ventas = {
    "Producto": ["Pan", "Leche", "Pan", "Queso", "Leche", "Bebida", "Pan"],
    "Cantidad": [10, 5, 8, 3, 7, 12, 6],
    "Precio Unitario": [1200, 1000, 1200, 3500, 1000, 1500, 1200]
}

df = pd.DataFrame(ventas)

archivo_excel = DATA_DIR / "ventas_cliente.xlsx"
archivo_csv = DATA_DIR / "ventas_cliente.csv"

df.to_excel(archivo_excel, index=False)
df.to_csv(archivo_csv, index=False)

print("Archivos de prueba creados correctamente.")
print(f"Excel: {archivo_excel}")
print(f"CSV: {archivo_csv}")
import pandas as pd

ventas = {
    "producto": ["Pan", "Leche", "Pan", "Queso", "Leche", "Bebida", "Pan"],
    "cantidad": [10, 5, 8, 3, 7, 12, 6],
    "precio_unitario": [1200, 1000, 1200, 3500, 1000, 1500, 1200]
}

df = pd.DataFrame(ventas)

df.to_excel("ventas_cliente.xlsx", index=False)

print("Archivo ventas_cliente.xlsx creado correctamente.")
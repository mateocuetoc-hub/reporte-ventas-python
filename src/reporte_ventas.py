from pathlib import Path
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

ARCHIVO_ENTRADA = DATA_DIR / "ventas_cliente.xlsx"
ARCHIVO_SALIDA = OUTPUT_DIR / "reporte_final.xlsx"

COLUMNAS_NECESARIAS = ["producto", "cantidad", "precio_unitario"]


def validar_archivo(df):
    df.columns = [str(col).strip().lower() for col in df.columns]

    columnas_faltantes = []

    for columna in COLUMNAS_NECESARIAS:
        if columna not in df.columns:
            columnas_faltantes.append(columna)

    if len(columnas_faltantes) > 0:
        print("Error: faltan columnas en el Excel de entrada.")
        print("Columnas faltantes:")
        for columna in columnas_faltantes:
            print(f"- {columna}")
        sys.exit(1)

    if df.empty:
        print("Error: el archivo de ventas está vacío.")
        sys.exit(1)

    df["producto"] = df["producto"].astype(str).str.strip()
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce")
    df["precio_unitario"] = pd.to_numeric(df["precio_unitario"], errors="coerce")

    if df["cantidad"].isnull().any():
        print("Error: hay valores inválidos en la columna 'cantidad'.")
        sys.exit(1)

    if df["precio_unitario"].isnull().any():
        print("Error: hay valores inválidos en la columna 'precio_unitario'.")
        sys.exit(1)

    if (df["cantidad"] <= 0).any():
        print("Error: la cantidad debe ser mayor a 0.")
        sys.exit(1)

    if (df["precio_unitario"] <= 0).any():
        print("Error: el precio unitario debe ser mayor a 0.")
        sys.exit(1)

    return df


def ajustar_columnas(ws):
    for columna in ws.columns:
        largo_maximo = 0
        letra_columna = get_column_letter(columna[0].column)

        for celda in columna:
            if celda.value is not None:
                largo_maximo = max(largo_maximo, len(str(celda.value)))

        ws.column_dimensions[letra_columna].width = largo_maximo + 4


def formatear_tabla(ws):
    color_encabezado = "1F4E78"
    borde_fino = Side(style="thin", color="D9E2F3")

    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=color_encabezado)
        celda.alignment = Alignment(horizontal="center")
        celda.border = Border(bottom=borde_fino)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for fila in ws.iter_rows(min_row=2):
        for celda in fila:
            encabezado = ws.cell(row=1, column=celda.column).value

            if encabezado in ["precio_unitario", "total"]:
                celda.number_format = '$#,##0'

            if encabezado == "cantidad":
                celda.number_format = '#,##0'

    ajustar_columnas(ws)


def crear_resumen_ejecutivo(wb, total_vendido, producto_mas_vendido, cantidad_mas_vendida):
    if "Resumen Ejecutivo" in wb.sheetnames:
        del wb["Resumen Ejecutivo"]

    ws = wb.create_sheet("Resumen Ejecutivo", 0)

    ws["A1"] = "REPORTE DE VENTAS"
    ws["A1"].font = Font(bold=True, size=18, color="1F4E78")

    ws["A3"] = "Total vendido"
    ws["B3"] = total_vendido
    ws["B3"].number_format = '$#,##0'

    ws["A4"] = "Producto más vendido"
    ws["B4"] = producto_mas_vendido

    ws["A5"] = "Cantidad vendida"
    ws["B5"] = cantidad_mas_vendida

    for celda in ["A3", "A4", "A5"]:
        ws[celda].font = Font(bold=True)

    for celda in ["A3", "A4", "A5", "B3", "B4", "B5"]:
        ws[celda].alignment = Alignment(horizontal="left")

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 25

    ws_resumen = wb["Resumen"]

    grafico = BarChart()
    grafico.title = "Ventas por producto"
    grafico.y_axis.title = "Total vendido"
    grafico.x_axis.title = "Producto"

    datos = Reference(
        ws_resumen,
        min_col=3,
        min_row=1,
        max_row=ws_resumen.max_row
    )

    categorias = Reference(
        ws_resumen,
        min_col=1,
        min_row=2,
        max_row=ws_resumen.max_row
    )

    grafico.add_data(datos, titles_from_data=True)
    grafico.set_categories(categorias)
    grafico.height = 8
    grafico.width = 16

    ws.add_chart(grafico, "D3")


def generar_reporte():
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not ARCHIVO_ENTRADA.exists():
        print("Error: no se encontró el archivo de entrada.")
        print(f"Debes tener este archivo: {ARCHIVO_ENTRADA}")
        sys.exit(1)

    df = pd.read_excel(ARCHIVO_ENTRADA)
    df = validar_archivo(df)

    df["total"] = df["cantidad"] * df["precio_unitario"]

    resumen_productos = df.groupby("producto").agg({
        "cantidad": "sum",
        "total": "sum"
    }).reset_index()

    total_vendido = df["total"].sum()

    producto_mas_vendido = resumen_productos.loc[
        resumen_productos["cantidad"].idxmax(),
        "producto"
    ]

    cantidad_mas_vendida = resumen_productos["cantidad"].max()

    try:
        with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Ventas", index=False)
            resumen_productos.to_excel(writer, sheet_name="Resumen", index=False)

        wb = load_workbook(ARCHIVO_SALIDA)

        formatear_tabla(wb["Ventas"])
        formatear_tabla(wb["Resumen"])

        crear_resumen_ejecutivo(
            wb,
            total_vendido,
            producto_mas_vendido,
            cantidad_mas_vendida
        )

        wb.save(ARCHIVO_SALIDA)

    except PermissionError:
        print("Error: no se pudo guardar el reporte.")
        print("Cierra el archivo reporte_final.xlsx si está abierto en LibreOffice.")
        sys.exit(1)

    print("Reporte generado correctamente.")
    print(f"Total vendido: ${total_vendido:,.0f}")
    print(f"Producto más vendido: {producto_mas_vendido} ({cantidad_mas_vendida} unidades)")
    print(f"Archivo creado: {ARCHIVO_SALIDA}")


if __name__ == "__main__":
    generar_reporte()
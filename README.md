# Automatizador de Reportes de Ventas

Este proyecto genera automáticamente reportes de ventas en Excel a partir de un archivo de entrada.

El sistema permite leer una planilla de ventas, calcular totales, generar un resumen por producto, crear un resumen ejecutivo y exportar un archivo Excel final con formato profesional, gráfico incluido y hoja de metadatos.

## ¿Qué hace?

El programa toma un archivo de ventas en formato Excel o CSV y genera un reporte final con:

- Total vendido
- Producto más vendido
- Cantidad vendida del producto principal
- Detalle de ventas
- Resumen por producto
- Gráfico de ventas por producto
- Nombre del negocio
- Fecha automática del reporte
- Hoja de metadatos
- Formato de moneda
- Archivo Excel final listo para revisar o entregar

## Estructura del proyecto

```text
reporte-ventas-python/
├── data/
│   ├── ventas_cliente.xlsx
│   └── ventas_cliente.csv
├── output/
│   └── reporte_final.xlsx
├── src/
│   ├── crear_ventas_cliente.py
│   └── reporte_ventas.py
├── README.md
├── requirements.txt
└── .gitignore
```

## Archivo de entrada

Por defecto, el programa busca el archivo:

```text
data/ventas_cliente.xlsx
```

También permite usar archivos:

```text
.xlsx
.xls
.csv
```

El archivo debe contener las siguientes columnas:

| producto | cantidad | precio_unitario |
|---|---:|---:|
| Pan | 10 | 1200 |
| Leche | 5 | 1000 |
| Queso | 3 | 3500 |

También acepta nombres de columnas con mayúsculas o espacios, por ejemplo:

```text
Producto
Cantidad
Precio Unitario
```

El programa normaliza los nombres de columnas para hacer el proceso más flexible.

## Archivo de salida

Por defecto, el programa genera:

```text
output/reporte_final.xlsx
```

El reporte generado incluye cuatro hojas:

- Resumen Ejecutivo
- Ventas
- Resumen
- Metadatos

## Tecnologías usadas

- Python
- Pandas
- OpenPyXL
- Excel / LibreOffice Calc

## Instalación

Crear entorno virtual:

```bash
python3 -m venv venv
```

Activar entorno virtual:

```bash
source venv/bin/activate
```

Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Uso

Crear archivos de prueba en Excel y CSV:

```bash
python src/crear_ventas_cliente.py
```

Generar reporte usando el archivo por defecto:

```bash
python src/reporte_ventas.py
```

Generar reporte indicando archivo de entrada:

```bash
python src/reporte_ventas.py data/ventas_cliente.xlsx
```

Generar reporte desde un archivo CSV:

```bash
python src/reporte_ventas.py data/ventas_cliente.csv
```

Generar reporte indicando archivo de entrada y nombre de salida personalizado:

```bash
python src/reporte_ventas.py data/ventas_cliente.xlsx -o output/reporte_junio.xlsx
```

Generar reporte indicando el nombre del negocio:

```bash
python src/reporte_ventas.py data/ventas_cliente.xlsx -o output/reporte_junio.xlsx -n "Almacén Don Mateo"
```

Generar reporte desde CSV con nombre de negocio:

```bash
python src/reporte_ventas.py data/ventas_cliente.csv -o output/reporte_csv.xlsx -n "Almacén Don Mateo"
```

Abrir el reporte generado:

```bash
libreoffice output/reporte_final.xlsx
```

También se puede abrir con:

```bash
xdg-open output/reporte_final.xlsx
```

## Parámetros disponibles

| Parámetro | Descripción |
|---|---|
| `entrada` | Ruta del archivo de entrada. Puede ser `.xlsx`, `.xls` o `.csv`. |
| `-o`, `--output` | Ruta y nombre del archivo Excel de salida. |
| `-n`, `--negocio` | Nombre del negocio que aparecerá en el reporte. |
| `--sep` | Separador para archivos CSV. Por defecto usa coma `,`. |

## Ejemplo de uso real

Un negocio puede entregar una planilla con sus ventas en Excel o CSV.

El programa procesa automáticamente esa información y genera un reporte ordenado con totales, resumen, gráfico, fecha del reporte y metadatos.

Ejemplo con Excel:

```bash
python src/reporte_ventas.py data/ventas_cliente.xlsx -o output/reporte_mensual.xlsx -n "Almacén Don Mateo"
```

Ejemplo con CSV:

```bash
python src/reporte_ventas.py data/ventas_cliente.csv -o output/reporte_mensual_csv.xlsx -n "Almacén Don Mateo"
```

Resultado:

```text
output/reporte_mensual.xlsx
```

## Validaciones incluidas

El programa valida que:

- El archivo de entrada exista
- El formato sea compatible
- El archivo no esté vacío
- Existan las columnas necesarias
- Las cantidades sean numéricas
- Los precios sean numéricos
- Las cantidades sean mayores a 0
- Los precios sean mayores a 0
- Los productos tengan nombre

## Estado del proyecto

Versión actual: 4

Características principales:

- Lectura de archivo Excel
- Lectura de archivo CSV
- Validación de columnas necesarias
- Validación de datos numéricos
- Generación de reporte Excel
- Formato profesional
- Gráfico de ventas por producto
- Soporte para archivo de entrada por parámetro
- Soporte para archivo de salida personalizado
- Soporte para nombre del negocio
- Fecha automática del reporte
- Hoja de metadatos
- Resumen ordenado por total vendido

## Objetivo

Este proyecto fue creado como una herramienta de automatización para reportes simples de ventas, pensada para pequeños negocios, emprendimientos o casos de práctica con Python y análisis de datos.


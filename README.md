# Automatizador de Reportes de Ventas

Este proyecto genera automáticamente reportes de ventas en Excel a partir de un archivo de entrada.

El sistema permite leer una planilla de ventas, calcular totales, generar un resumen por producto, crear un resumen ejecutivo y exportar un archivo Excel final con formato profesional y gráfico incluido.

## ¿Qué hace?

El programa toma un archivo Excel con ventas y genera un reporte final con:

- Total vendido
- Producto más vendido
- Cantidad vendida del producto principal
- Detalle de ventas
- Resumen por producto
- Gráfico de ventas por producto
- Formato de moneda
- Archivo Excel final listo para revisar o entregar

## Estructura del proyecto

```text
reporte-ventas-python/
├── data/
│   └── ventas_cliente.xlsx
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

El reporte generado incluye tres hojas:

- Resumen Ejecutivo
- Ventas
- Resumen

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

Crear archivo de prueba:

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

Generar reporte indicando archivo de entrada y nombre de salida personalizado:

```bash
python src/reporte_ventas.py data/ventas_cliente.xlsx -o output/reporte_junio.xlsx
```

Abrir el reporte generado:

```bash
libreoffice output/reporte_final.xlsx
```

También se puede abrir con:

```bash
xdg-open output/reporte_final.xlsx
```

## Ejemplo de uso real

Un negocio puede entregar una planilla con sus ventas en Excel.  
El programa procesa automáticamente esa información y genera un reporte ordenado con totales, resumen y gráfico.

Ejemplo:

```bash
python src/reporte_ventas.py data/ventas_cliente.xlsx -o output/reporte_mensual.xlsx
```

Resultado:

```text
output/reporte_mensual.xlsx
```

## Estado del proyecto

Versión actual: 3

Características principales:

- Lectura de archivo Excel
- Validación de columnas necesarias
- Validación de datos numéricos
- Generación de reporte Excel
- Formato profesional
- Gráfico de ventas por producto
- Soporte para archivo de entrada por parámetro
- Soporte para archivo de salida personalizado

## Objetivo

Este proyecto fue creado como una herramienta de automatización para reportes simples de ventas, pensada para pequeños negocios, emprendimientos o casos de práctica con Python y análisis de datos.
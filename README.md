# Automatizador de Reportes de Ventas

Este proyecto genera automáticamente un reporte de ventas en Excel a partir de un archivo de entrada.

## ¿Qué hace?

El programa lee un archivo llamado `ventas_cliente.xlsx`, calcula el total vendido por producto y genera un archivo final llamado `reporte_final.xlsx`.

## Archivo de entrada

El archivo `ventas_cliente.xlsx` debe tener las siguientes columnas:

- producto
- cantidad
- precio_unitario

Ejemplo:

| producto | cantidad | precio_unitario |
|---|---:|---:|
| Pan | 10 | 1200 |
| Leche | 5 | 1000 |
| Queso | 3 | 3500 |

## Archivo de salida

El programa genera `reporte_final.xlsx` con tres hojas:

- Ventas
- Resumen
- Resumen Ejecutivo

## Tecnologías usadas

- Python
- Pandas
- OpenPyXL
- Excel / LibreOffice Calc

## Cómo ejecutar

Primero activar el entorno virtual:

```bash
source venv/bin/activate
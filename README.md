# Automatizador de Reportes de Ventas

Este proyecto genera automáticamente un reporte de ventas en Excel a partir de un archivo de entrada.

## ¿Qué hace?

El sistema lee un archivo `ventas_cliente.xlsx`, calcula totales de ventas, resume los datos por producto y genera un archivo `reporte_final.xlsx` con formato profesional.

## Estructura del proyecto

```text
reporte_ventas/
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

El archivo debe estar en la carpeta `data/` y llamarse:

```text
ventas_cliente.xlsx
```

Debe tener las siguientes columnas:

| producto | cantidad | precio_unitario |
|---|---:|---:|
| Pan | 10 | 1200 |
| Leche | 5 | 1000 |
| Queso | 3 | 3500 |

## Archivo de salida

El programa genera el archivo:

```text
output/reporte_final.xlsx
```

El reporte incluye:

- Resumen ejecutivo
- Detalle de ventas
- Resumen por producto
- Formato de moneda
- Gráfico de ventas por producto

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

Generar reporte:

```bash
python src/reporte_ventas.py
```

Abrir reporte:

```bash
libreoffice output/reporte_final.xlsx
```
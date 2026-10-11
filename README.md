# 📊 Lector de Excel y Generador de Reporte de Ventas

Script en Python que lee un archivo Excel con datos de ventas, calcula totales por vendedor y genera un reporte formateado en un nuevo archivo Excel.

##  Características

- Lectura de archivos Excel con **pandas**.
- Validación automática de columnas requeridas.
- Cálculo de totales por vendedor (total vendido, cantidad de productos, número de ventas).
- Exportación a Excel con **openpyxl** y ajuste automático del ancho de columnas.
- Código organizado con **Programación Orientada a Objetos (POO)**.
- Manejo de errores con `try/except`.

##  Tecnologías utilizadas

- Python 3.10+
- pandas
- openpyxl

## Requisitos previos

- Tener Python instalado.
- Tener el archivo `ventas.xlsx` en la misma carpeta del script (o modificar la ruta en el código).

##  Instalación

1. Clona el repositorio:
   ```bash
   git clone https://github.com/TU_USUARIO/reporte-ventas-python.git
   cd reporte-ventas-python

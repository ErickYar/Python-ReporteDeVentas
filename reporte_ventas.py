"""
Proyecto: Lector de Excel y generador de reporte de ventas
Autor: Erick Yarleque
Descripción: Lee un archivo Excel con ventas, calcula totales por vendedor
y genera un reporte en un nuevo archivo Excel.
"""

import pandas as pd
from datetime import datetime


class ReporteVentas:
    """Clase que procesa un archivo Excel de ventas y genera un reporte."""

    def __init__(self, archivo_entrada):
        self.archivo_entrada = archivo_entrada
        self.datos = None
        self.resumen = None

    def leer_excel(self):
        """Lee el archivo Excel de entrada."""
        try:
            self.datos = pd.read_excel(self.archivo_entrada)
            print(f"✅ Archivo leído: {len(self.datos)} filas encontradas.")
        except FileNotFoundError:
            print(f"❌ No se encontró el archivo: {self.archivo_entrada}")
            raise

    def validar_columnas(self):
        """Verifica que el Excel tenga las columnas necesarias."""
        columnas_necesarias = {"Vendedor", "Producto", "Cantidad", "Precio"}
        if not columnas_necesarias.issubset(self.datos.columns):
            faltantes = columnas_necesarias - set(self.datos.columns)
            raise ValueError(f"Faltan columnas en el Excel: {faltantes}")

    def calcular_resumen(self):
        """Calcula el total vendido por vendedor."""
        self.datos["Total"] = self.datos["Cantidad"] * self.datos["Precio"]
        self.resumen = (
            self.datos.groupby("Vendedor")
            .agg(
                Total_Vendido=("Total", "sum"),
                Cantidad_Productos=("Cantidad", "sum"),
                Num_Ventas=("Total", "count"),
            )
            .reset_index()
            .sort_values("Total_Vendido", ascending=False)
        )
        print("✅ Resumen calculado por vendedor.")

    def generar_reporte(self, archivo_salida="reporte_ventas.xlsx"):
        """Exporta el resumen a un nuevo archivo Excel con formato."""
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
        with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
            self.resumen.to_excel(writer, sheet_name="Resumen", index=False)

            # Ajustar ancho de columnas
            hoja = writer.sheets["Resumen"]
            for columna in hoja.columns:
                largo_max = max(len(str(celda.value)) for celda in columna if celda.value)
                hoja.column_dimensions[columna[0].column_letter].width = largo_max + 4

        print(f"✅ Reporte generado: {archivo_salida}")
        print(f"   Fecha de generación: {fecha}")


def main():
    reporte = ReporteVentas("ventas.xlsx")
    reporte.leer_excel()
    reporte.validar_columnas()
    reporte.calcular_resumen()
    reporte.generar_reporte()

    print("\n📊 Vista previa del resumen:")
    print(reporte.resumen.to_string(index=False))


if __name__ == "__main__":
    main()
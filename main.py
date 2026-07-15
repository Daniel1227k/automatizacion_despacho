import pandas as pd
from docxtpl import DocxTemplate
from pathlib import Path
import re


ARCHIVO_EXCEL = "prueba_vf.xlsx"
PLANTILLA = "ejemplo.docx"

CARPETA_SALIDA = Path("Reportes_Finales")
CARPETA_SALIDA.mkdir(exist_ok=True)

# ===========================
# FUNCIONES
# ===========================

def limpiarNombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre))

def esSucursal(nombre):
    #Devuelve True si el nombre corresponde
    # a una sucursal.
    if pd.isna(nombre):
        return False

    nombre = str(nombre).strip().upper()

    prefijos = (
        "SUC.",
        "MTZ.",
        "CENTRAL"
    )
    return nombre.startswith(prefijos)

def leerClientes():
    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name="CLIENTES",
        header=2
        )

    df.dropna(
        subset=["NOMBRE DEL CLIENTE"]
    )

    return df

clientes = leerClientes()

print(clientes.head())

def leerSucursales():
    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name="SUCURSALES",
        header=2
    )

    df = df.dropna(
        subset=["SUCURSALES"],
        how="all"
    )

    return df

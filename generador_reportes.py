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
    """Devuelve True si el nombre corresponde a una sucursal."""
    if pd.isna(nombre):
        return False
    nombre = str(nombre).strip().upper()
    prefijos = ("SUC.", "MTZ.", "CENTRAL")
    return nombre.startswith(prefijos)


def leerClientes():

    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name="CLIENTES",
        header=[2, 3]
    )

    cols = list(df.columns)
    cols[1] = ("NOMBRE DEL CLIENTE", "")
    df.columns = pd.MultiIndex.from_tuples(cols)

    df = df.dropna(subset=[("NOMBRE DEL CLIENTE", "")])

    return df


def leerSucursales():
    # Mismo tratamiento que leerClientes(): header=[2, 3] para no perder
    # el texto específico de cada trámite en las columnas de STPS.
    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name="SUCURSALES",
        header=[2, 3]
    )

    cols = list(df.columns)
    cols[1] = ("SUCURSALES", "")
    df.columns = pd.MultiIndex.from_tuples(cols)

    # Esta reasignación ya la tenías bien.
    df = df.dropna(subset=[("SUCURSALES", "")], how="all")

    return df


def generarWord():
    # PENDIENTE: todavía no la escribo — depende de cómo resolvamos el punto
    # de las sucursales con datos propios en su fila de cliente (ver chat).
    # En cuanto lo confirmes, seguimos aquí.
    pass


if __name__ == "__main__":
    clientes = leerClientes()
    sucursales = leerSucursales()
    print(f"CLIENTES: {len(clientes)} filas")
    print(clientes.head())
    print()
    print(f"SUCURSALES: {len(sucursales)} filas")
    print(sucursales[("SUCURSALES", "")].tolist())

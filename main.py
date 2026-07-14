import pandas as pd
from docxtpl import DocxTemplate
import os

# CONFIGURACIÓN
archivo_excel = "prueba_vf.xlsx"
archivo_plantilla = "plantilla_maestra.docx"
os.makedirs("Reportes_Finales", exist_ok=True)

def es_sucursal(nombre):
    """
    Función que decide si algo es una sucursal basándose en su nombre.
    """
    if pd.isna(nombre): return False
    nombre = str(nombre).upper()
    prefijos = ["SUC.", "MTZ.", "CENTRAL"]
    return any(nombre.startswith(p) for p in prefijos)

def generar_word(datos, tiene_sucursales=False, lista_sucursales=None):
    doc = DocxTemplate(archivo_plantilla)
    contexto = datos.copy()
    contexto['tiene_sucursales'] = tiene_sucursales
    if tiene_sucursales:
        contexto['lista_sucursales'] = lista_sucursales
    
    nombre_archivo = f"Reportes_Finales/Reporte_{datos['NOMBRE DEL CLIENTE']}.docx"
    doc.render(contexto)
    doc.save(nombre_archivo)
    print(f"Generado: {nombre_archivo}")

# --- Hoja 1: CLIENTES SIN SUCURSALES (Hoja CLIENTES) ---
print("Procesando Clientes hoja 'CLIENTES'...")
df_clientes = pd.read_excel(archivo_excel, sheet_name="CLIENTES", header=2)
df_clientes = df_clientes.dropna(subset=['NOMBRE DEL CLIENTE'])

for index, fila in df_clientes.iterrows():
    generar_word(fila.to_dict(), tiene_sucursales=False)

# --- Hoja 2: CLIENTES CON SUCURSALES (Hoja SUCURSALES) ---
print("Procesando Sucursales hoja 'SUCURSALES'...")
df_suc = pd.read_excel(archivo_excel, sheet_name="SUCURSALES", header=2)

cliente_actual = None
lista_sucursales_temp = []
datos_cliente_actual = None

for index, fila in df_suc.iterrows():
    nombre_en_celda = fila['SUCURSALES']
    
    # Si no es nulo y es SUCURSAL (usamos nuestra función inteligente)
    if pd.notna(nombre_en_celda) and es_sucursal(nombre_en_celda):
        lista_sucursales_temp.append(fila.to_dict())
    
    # Si NO es sucursal, asumimos que es un CLIENTE NUEVO
    elif pd.notna(nombre_en_celda):
        # Primero guardamos el cliente anterior si existía
        if cliente_actual is not None:
            generar_word(datos_cliente_actual, tiene_sucursales=True, lista_sucursales=lista_sucursales_temp)
        
        # Iniciamos nuevo cliente
        cliente_actual = nombre_en_celda
        datos_cliente_actual = {'NOMBRE DEL CLIENTE': cliente_actual}
        lista_sucursales_temp = []

# Generar el último cliente que quedó pendiente al final del ciclo
if cliente_actual is not None:
    generar_word(datos_cliente_actual, tiene_sucursales=True, lista_sucursales=lista_sucursales_temp)

print("¡Proceso terminado con éxito! Revisa la carpeta 'Reportes_Finales'.")
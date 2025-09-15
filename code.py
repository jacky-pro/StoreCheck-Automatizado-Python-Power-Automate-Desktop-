import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl as xl
import xlsxwriter
import os
import sys
 
directorio = os.path.dirname(os.path.abspath(sys.argv[0]))
 
base_file = os.path.join(directorio, "base.xlsx")
# base_crono_file = os.path.join(directorio, "base_crono.xlsx")
 
# Verificar que los archivos existan
if not os.path.exists(base_file):
    print(f"Error: No se encuentra el archivo {base_file}")
    sys.exit()
 
# if not os.path.exists(base_crono_file):
#     print(f"Error: No se encuentra el archivo {base_crono_file}")
#     sys.exit()
 
# cargar base
try:
    df = pd.read_excel(base_file)
    print(f"Archivo base.xlsx cargado exitosamente. Filas: {len(df)}")
except Exception as e:
    print(f"Error al cargar base.xlsx: {e}")
    sys.exit()
 
# subir base crono
# try:
#     base_crono = pd.read_excel(base_crono_file)
#     print(f"Archivo base_crono.xlsx cargado exitosamente. Filas: {len(base_crono)}")
# except Exception as e:
#     print(f"Error al cargar base_crono.xlsx: {e}")
#     sys.exit()
 
# filtrar por actividad
columna = "Act. Promocional"
 
# filtrar valores unicos de la actividad
valores = df[columna].dropna().unique()
 
# crear diccionario para almacenar datos de la actividad seleccionada
resultado = {}
 
 
# crear dataframe con información de la actividad elegida
def filtrar():
    valor_elegido = combo.get()
    if valor_elegido:
        resultado["df_filtrado"] = df[df[columna] == valor_elegido]
        resultado["promocion_elegida"] = valor_elegido  # Guardar el nombre de la promoción
    root.destroy()
 
 
# Crear ventana
root = tk.Tk()
root.title("Seleccionar valor para filtrar")
root.geometry("400x200")

label = tk.Label(root, text=f"Selecciona un valor de '{columna}':")
label.pack(pady=10)

combo = ttk.Combobox(root, values=list(valores), state="readonly")
combo.pack(padx=40)

boton = tk.Button(root, text="Aceptar", command=filtrar)
boton.pack(pady=10)

# Centrar la ventana
root.update_idletasks()
x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
root.geometry(f"+{x}+{y}")

root.mainloop()
 
# guardar diccionario en una dataframe y verificar seleccion
df_resultado = resultado.get("df_filtrado")
if df_resultado is None:
    messagebox.showinfo("Información", "No se seleccionó ninguna actividad promocional. El programa terminará.")
    sys.exit()
 
# elegir los puntos de venta sin actividad
 
# obtener puntos de venta sin duplicados
pdv_sin_dup = df_resultado["PDV_nombre"].dropna().unique()
 
# Crear ventana
root = tk.Tk()
root.title("Selecciona puntos de venta sin actividad para aplicar 'SA'")
root.geometry("500x400")

# Instrucción
label = tk.Label(root, text="Selecciona los puntos de venta sin actividad:")
label.pack(pady=10)

# Listbox con selección múltiple
lista = tk.Listbox(root, selectmode=tk.MULTIPLE, height=10)
for punto in pdv_sin_dup:
    lista.insert(tk.END, punto)
lista.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)

# Diccionario para almacenar resultado
resultado_pdv = {}


def filtrar_pdv():
    seleccionados = [lista.get(i) for i in lista.curselection()]
    if not seleccionados:
        messagebox.showwarning("Aviso", "Debes seleccionar al menos un punto de venta.")
        return
    # Filtrar solo los seleccionados
    df_filtrado = df_resultado[df_resultado["PDV_nombre"].isin(seleccionados)]
    resultado_pdv["df"] = df_filtrado
    root.destroy()


# Botón
boton = tk.Button(root, text="Filtrar", command=filtrar_pdv)
boton.pack(pady=10)

# Centrar la ventana
root.update_idletasks()
x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
root.geometry(f"+{x}+{y}")

root.mainloop()
 
# Mostrar el nuevo DataFrame
df_modificado = resultado_pdv.get("df")
 
# Verificar si se seleccionaron puntos de venta
if df_modificado is None:
    print("No se seleccionaron puntos de venta. El programa terminará.")
    sys.exit()
 
print(f"Puntos de venta seleccionados: {len(df_modificado)} filas")
 
# eliminar valores nulos dentro de la columna PRODUCTO
df_modificado_v2 = df_modificado[df_modificado["PRODUCTO"].notna()]
print(f"Después de eliminar PRODUCTO nulos: {len(df_modificado_v2)} filas")
 
# quedarme con las columnas que solo voy a utilizar
df_modificado_v3 = df_modificado_v2[
    ["PDV_nombre", "FECHA", "PUESTO DE MERCADO", "Act. Promocional", "PRODUCTO", "STOCK FINAL", "STOCK INICIAL"]]
print(f"Columnas seleccionadas: {list(df_modificado_v3.columns)}")
 
# hacer el groupby
try:
    df_modificado_v4 = \
        df_modificado_v3.groupby(["PDV_nombre", "FECHA", "PUESTO DE MERCADO", "Act. Promocional", "PRODUCTO"])[
            ["STOCK INICIAL", "STOCK FINAL"]].sum().reset_index()
    print(f"Después del groupby: {len(df_modificado_v4)} filas")
except Exception as e:
    print(f"Error en el groupby: {e}")
    # Intentar con solo las columnas numéricas
    df_modificado_v4 = \
        df_modificado_v3.groupby(["PDV_nombre", "FECHA", "PUESTO DE MERCADO", "Act. Promocional", "PRODUCTO"])[
            ["STOCK INICIAL", "STOCK FINAL"]].agg('sum').reset_index()
    print(f"Después del groupby (método alternativo): {len(df_modificado_v4)} filas")
 
# Mantener PUESTO DE MERCADO como está (sin separar) y renombrarlo a NOMBRE CLIENTE
df_modificado_v5 = df_modificado_v4.copy()
df_modificado_v5 = df_modificado_v5.rename(columns={'PUESTO DE MERCADO': 'NOMBRE CLIENTE'})
print(f"Columna renombrada: {list(df_modificado_v5.columns)}")
 
# cruzar dex y provincia
# base_crono["PDV_nombre"] = base_crono["MERCADO"]
# print(f"Antes del merge: {len(df_modificado_v5)} filas en df_modificado_v5, {len(base_crono)} filas en base_crono")
 
# df_modificado_v6 = pd.merge(df_modificado_v5, base_crono, on='PDV_nombre', how='inner')
# print(f"Después del merge: {len(df_modificado_v6)} filas")
 
# reordenar las columnas
# df_modificado_v7 = df_modificado_v6[
#     ["FECHA", "PROVINCIA", "DEX", "MERCADO", "NOMBRE CLIENTE", "PRODUCTO", "STOCK INICIAL", "STOCK FINAL"]]
# print(f"Columnas finales: {list(df_modificado_v7.columns)}")
 
# Usar df_modificado_v5 directamente ya que no tenemos base_crono
df_modificado_v7 = df_modificado_v5.copy()
# Agregar columnas faltantes con valores por defecto
df_modificado_v7["PROVINCIA"] = "N/A"
df_modificado_v7["DEX"] = "N/A"
df_modificado_v7["MERCADO"] = df_modificado_v7["PDV_nombre"]
print(f"Columnas finales: {list(df_modificado_v7.columns)}")
 
# In[69]:
 
 
# hacer la tabla dinámica , eliminar los ceros
try:
    pv = pd.pivot_table(
        data=df_modificado_v7,
        index=["FECHA", "PROVINCIA", "DEX", 'MERCADO', 'NOMBRE CLIENTE'],  # Lo que va como filas
        columns='PRODUCTO',  # Lo que va como columnas
        values=["STOCK INICIAL", 'STOCK FINAL'],  # Lo que va dentro de la tabla (valores)
        aggfunc='sum',
        fill_value=0
    )
except Exception as e:
    print(f"Error al crear tabla dinámica: {e}")
    # Crear tabla dinámica simplificada
    pv = pd.pivot_table(
        data=df_modificado_v7,
        index=["FECHA", "PROVINCIA", "DEX", 'MERCADO', 'NOMBRE CLIENTE'],
        columns='PRODUCTO',
        values=["STOCK INICIAL", 'STOCK FINAL'],
        aggfunc='sum',
        fill_value=0
    )
 
# --- Reordenar columnas para que STOCK INICIAL vaya primero ---
if isinstance(pv.columns, pd.MultiIndex):
    # Obtener todas las tuplas de las columnas (ej: ('STOCK INICIAL', 'Producto A'))
    cols = pv.columns.to_list()
   
    # Definir el orden deseado para el primer nivel del encabezado
    level0_order = ['STOCK INICIAL', 'STOCK FINAL']
   
    # Ordenar las columnas: primero por el orden de level0_order, luego alfabéticamente por el nombre del producto
    sorted_cols = sorted(cols, key=lambda c: (level0_order.index(c[0]), c[1]))
   
    # Reorganizar el DataFrame usando la lista de columnas ordenada
    pv = pv[sorted_cols]
 
pv = pv.reset_index()
 
cols_numericas = pv.select_dtypes(include='number').columns
 
# Reemplaza ceros por NaN
pv[cols_numericas] = pv[cols_numericas].replace(0, np.nan)
 
# Ordenar por mercado y luego por cliente para asegurar una numeración correcta y ascendente
pv = pv.sort_values(by=['MERCADO', 'NOMBRE CLIENTE']).reset_index(drop=True)
 
# Agregar columna N° con numeración por mercado
pv = pv.sort_index()  # Ordenar el índice para evitar warnings de rendimiento
numeration = pv.groupby('MERCADO').cumcount() + 1
pv.insert(0, 'N°', numeration)
 
# In[70]:
 
 
###### HOJA RESUMEN
 
df_modificado_v8 = df_modificado_v7.copy()
 
# eliminar ceros
df_modificado_v8["STOCK INICIAL"] = df_modificado_v8["STOCK INICIAL"].replace(0, np.nan)
df_modificado_v8["STOCK FINAL"] = df_modificado_v8["STOCK FINAL"].replace(0, np.nan)
 
# In[71]:
 
 
# hacer recuento de puestos de mercado por mercado
 
df_modificado_v8_drop = df_modificado_v8[["MERCADO", "PROVINCIA", "NOMBRE CLIENTE"]].drop_duplicates()
 
df_modificado_v10 = df_modificado_v8_drop.groupby(["PROVINCIA", "MERCADO"]).size().reset_index(
    name='PUESTOS DE MERCADO')
 
# copiar el contenido a dos columnas
df_modificado_v10["PUESTOS ATENDIDOS POR ALICORP"] = df_modificado_v10["PUESTOS DE MERCADO"]
df_modificado_v10["PRESENCIA  DEL PRODUCTO"] = df_modificado_v10["PUESTOS DE MERCADO"]
 
df_modificado_v10["COBERTURA TOTAL (Puestos Atendidos por Alicorp)"] = '100%'
 
# stock final en RESUMEN
 
df_modificado_v11_drop = df_modificado_v8[["MERCADO", "PROVINCIA", "DEX", "NOMBRE CLIENTE"]].drop_duplicates()
df_modificado_v12 = df_modificado_v11_drop.groupby(["PROVINCIA", "MERCADO", "DEX"]).size().reset_index(
    name="PUESTOS DE MERCADO")
 
# calcular stock total
try:
    df_modificado_v8_stock_total = df_modificado_v8.groupby("MERCADO")["STOCK FINAL"].sum().reset_index()
except Exception as e:
    print(f"Error al calcular stock total: {e}")
    df_modificado_v8_stock_total = pd.DataFrame(columns=["MERCADO", "STOCK FINAL"])

# hacer join df_modificado_v12 y df_modificado_v8_stock_total
try:
    df_modificado_v13 = pd.merge(df_modificado_v8_stock_total, df_modificado_v12, on="MERCADO")
except Exception as e:
    print(f"Error al hacer merge: {e}")
    # Crear DataFrame con columnas por defecto
    df_modificado_v13 = pd.DataFrame(columns=["MERCADO", "STOCK FINAL", "PROVINCIA", "DEX", "PUESTOS DE MERCADO"])
 
# cambiar nombre de columna de STOCK FINAL a STOCK TOTAL
 
df_modificado_v13 = df_modificado_v13.rename(columns={"STOCK FINAL": "STOCK TOTAL"})
 
# añadir objetivo unidades
df_modificado_v13["OBJETIVO UNIDADES"] = 0  # Cambiado de string vacío a número
df_modificado_v13["STOCK VS OBJETIVO X MERCADO"] = df_modificado_v13["STOCK TOTAL"] - df_modificado_v13[
    "OBJETIVO UNIDADES"]
 
# reordenar las columnas
resumen_stock_final = df_modificado_v13[
    ["PROVINCIA", "MERCADO", "DEX", "PUESTOS DE MERCADO", "STOCK TOTAL", "OBJETIVO UNIDADES",
     "STOCK VS OBJETIVO X MERCADO"]]
 
# In[73]:
 
 
##### HOJA STORECHECK
 
# --- Construir el nombre de la columna dinámicamente ---
promocion_elegida = resultado.get("promocion_elegida", "Bolivar")  # Usar "Bolivar" como fallback
columna_promocion = f"Vende {promocion_elegida} ¿Si/NO?"
 
# copiar el dataframe anterior para crear STORECHECK
# añadir las otras columnas
df_modificado_v16 = df_modificado_v8.copy()
 
df_modificado_v16[columna_promocion] = "1"
df_modificado_v16["Alicorp"] = "1"
df_modificado_v16["Otros"] = ""
df_modificado_v16["Presencia en Puestos Alicorp"] = "1"
df_modificado_v16["Presencia Producto"] = "1"
 
# reemplazar los datos por unos (ya está validado que no hay ceros)
df_modificado_v17 = df_modificado_v16
df_modificado_v17["STOCK FINAL"] = df_modificado_v17["STOCK FINAL"].where(df_modificado_v17["STOCK FINAL"].isna(), 1)
 
# ordenar las columnas y quedarme con las que voy a utilizar
df_modificado_v18 = df_modificado_v17[
    ["FECHA", "PROVINCIA", "DEX", "MERCADO", "NOMBRE CLIENTE", columna_promocion, "Alicorp", "Otros",
     "Presencia en Puestos Alicorp", "Presencia Producto", "PRODUCTO", "STOCK FINAL"]]
 
# crear la tabla dinámica
try:
    pv_tb2 = pd.pivot_table(
        data=df_modificado_v18,
        index=["FECHA", "PROVINCIA", "DEX", 'MERCADO', 'NOMBRE CLIENTE', columna_promocion, "Alicorp", "Otros",
               "Presencia en Puestos Alicorp", "Presencia Producto"],  # Lo que va como filas
        columns='PRODUCTO',  # Lo que va como columnas
        values=['STOCK FINAL'],  # <-- CORRECCIÓN: Usar una lista para forzar un encabezado de dos niveles
        aggfunc='sum',
        fill_value=0
    ).reset_index()
except Exception as e:
    print(f"Error al crear tabla dinámica STORECHECK: {e}")
    # Crear tabla dinámica simplificada
    pv_tb2 = pd.pivot_table(
        data=df_modificado_v18,
        index=["FECHA", "PROVINCIA", "DEX", 'MERCADO', 'NOMBRE CLIENTE'],
        columns='PRODUCTO',
        values=['STOCK FINAL'],
        aggfunc='sum',
        fill_value=0
    ).reset_index()
 
cols_numericas_2 = pv_tb2.select_dtypes(include='number').columns
 
# Reemplaza ceros por NaN
pv_tb2[cols_numericas_2] = pv_tb2[cols_numericas_2].replace(0, np.nan)
 
# Ordenar por mercado y luego por cliente para asegurar una numeración correcta y ascendente
pv_tb2 = pv_tb2.sort_values(by=['MERCADO', 'NOMBRE CLIENTE']).reset_index(drop=True)
 
# Agregar columna N° con numeración por mercado
pv_tb2 = pv_tb2.sort_index()  # Ordenar el índice para evitar warnings de rendimiento
numeration_2 = pv_tb2.groupby('MERCADO').cumcount() + 1
pv_tb2.insert(0, 'N°', numeration_2)
 
# In[82]:
 
 
# definir los dataframes y a que hoja del reporte storecheck pertenecen
 
# HOJA RESUMEN
# df_modificado_v10
# resumen_stock_final
 
# HOJA STORECHECK
# pv_tb2
 
# HOJA STOCK SIN ACTIVIDAD
# pv
 
# exportar con esos nombre
# definir diccionario con nombres de hojas y dataframes
hojas = {
    "RESUMEN": df_modificado_v10,  # Solo la primera tabla, la segunda se agregará después
    "STORECHECK": pv_tb2,
    "STOCK SIN ACTIVIDAD": pv
}
 
# Verificar que todos los DataFrames existan y no estén vacíos
for nombre_hoja, df in hojas.items():
    if df is None:
        print(f"Error: DataFrame para hoja '{nombre_hoja}' es None")
        sys.exit()
    if len(df) == 0:
        print(f"Advertencia: DataFrame para hoja '{nombre_hoja}' está vacío")
 
# --- INICIO: Generar Excel con formato personalizado ---

output_file = os.path.join(directorio, 'storecheck_borrador.xlsx')
print(f"Generando archivo Excel con formato: {output_file}")

try:
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
 
        # --- Definir Formatos exactos como el archivo de referencia ---
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#2F5597',
            'font_color': 'white',
            'border': 1,
            'font_size': 10
        })
       
        rotated_header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#DDEBF7',
            'border': 1,
            'rotation': 90,
            'font_size': 9
        })
       
        date_format = workbook.add_format({
            'num_format': 'm/d/yyyy',
            'align': 'center',
            'border': 1,
            'valign': 'vcenter'
        })
       
        body_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 9
        })
       
        number_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0',
            'font_size': 9
        })
 
        # --- Escribir cada hoja ---
        for sheet_name, df in hojas.items():
            print(f"Procesando hoja: {sheet_name}")
            if df.empty:
                print(f"  Advertencia: La hoja '{sheet_name}' está vacía.")
                continue

            worksheet = workbook.add_worksheet(sheet_name)
            is_pivot = any(isinstance(c, tuple) for c in df.columns)

            # --- Configurar tamaños de filas y columnas ---
            worksheet.set_row(0, 25)  # Altura de encabezados
            worksheet.set_row(1, 60)  # Altura para encabezados rotados
           
            # --- Escribir Encabezados ---
            if not is_pivot:
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    # Ajustar ancho de columna según el contenido
                    col_width = max(len(str(value)) + 2, 10)
                    worksheet.set_column(col_num, col_num, col_width)
                header_offset = 1
            else:
                for col_num, col_name in enumerate(df.columns):
                    if isinstance(col_name, tuple) and len(col_name) > 1 and col_name[0] not in ['N°', 'FECHA', 'PROVINCIA', 'DEX', 'MERCADO', 'NOMBRE CLIENTE', 'Vende Bolivar ¿Si/NO?', 'Alicorp', 'Otros', 'Presencia en Puestos Alicorp', 'Presencia Producto']:
                        worksheet.write(0, col_num, col_name[0], header_format)
                        worksheet.write(1, col_num, col_name[1], rotated_header_format)
                        # Ajustar ancho para columnas rotadas
                        worksheet.set_column(col_num, col_num, 4)
                    else:
                        display_name = col_name[0] if isinstance(col_name, tuple) and len(col_name) > 0 else col_name
                        worksheet.merge_range(0, col_num, 1, col_num, display_name, header_format)
                        # Ajustar ancho según el contenido
                        col_width = max(len(str(display_name)) + 2, 12)
                        worksheet.set_column(col_num, col_num, col_width)
               
                # Manejar encabezados fusionados
                merged_headers = {}
                for col_num, col_name in enumerate(df.columns):
                    if isinstance(col_name, tuple) and len(col_name) > 0:
                        level0 = col_name[0]
                        if level0 not in merged_headers:
                            merged_headers[level0] = {'start': col_num, 'count': 1}
                        else:
                            merged_headers[level0]['count'] += 1
               
                for header, props in merged_headers.items():
                    if props['count'] > 1:
                        worksheet.merge_range(0, props['start'], 0, props['start'] + props['count'] - 1, header, header_format)
                header_offset = 2

            # --- Escribir Datos ---
            current_row = header_offset
            date_col_idx = -1
            for i, col in enumerate(df.columns):
                col_name = col[0] if isinstance(col, tuple) and len(col) > 0 else col
                if col_name == 'FECHA':
                    date_col_idx = i
                    break

            market_col = next((c for c in df.columns if (c[0] if isinstance(c, tuple) and len(c) > 0 else c) == 'MERCADO'), None)
           
            apply_spacing = sheet_name in ["STORECHECK", "STOCK SIN ACTIVIDAD"]
            data_to_write = df.groupby(market_col) if (market_col and apply_spacing) else [(None, df)]
           
            for _, group in data_to_write:
                for _, row_data in group.iterrows():
                    for col_num, cell_value in enumerate(row_data):
                        fmt = body_format
                        val = cell_value
                       
                        # Determinar el formato correcto según el tipo de dato
                        if col_num == date_col_idx and pd.notna(val):
                            if isinstance(val, pd.Timestamp):
                                fmt = date_format
                            elif isinstance(val, str):
                                try:
                                    # Intentar convertir string a datetime con formato específico
                                    val = pd.to_datetime(val, dayfirst=True)
                                    fmt = date_format
                                except:
                                    fmt = body_format
                            else:
                                fmt = body_format
                        elif pd.isna(val):
                            val = ''
                        elif isinstance(val, (int, float)) and val != '':
                            fmt = number_format
                       
                        worksheet.write(current_row, col_num, val, fmt)
                    current_row += 1
                if market_col and apply_spacing:
                    current_row += 3
           
            # Si es la hoja RESUMEN, agregar la segunda tabla después de un espacio
            if sheet_name == "RESUMEN":
                # Agregar espacio entre tablas
                current_row += 3
               
                # Escribir encabezados de la segunda tabla
                for col_num, value in enumerate(resumen_stock_final.columns.values):
                    worksheet.write(current_row, col_num, value, header_format)
                    # Ajustar ancho de columna
                    col_width = max(len(str(value)) + 2, 12)
                    worksheet.set_column(col_num, col_num, col_width)
                current_row += 1
               
                # Escribir datos de la segunda tabla
                for _, row_data in resumen_stock_final.iterrows():
                    for col_num, cell_value in enumerate(row_data):
                        fmt = body_format
                        val = cell_value
                        if pd.isna(val):
                            val = ''
                        elif isinstance(val, (int, float)) and val != '':
                            fmt = number_format
                        worksheet.write(current_row, col_num, val, fmt)
                    current_row += 1
 
    print(f"\n¡Archivo Excel generado exitosamente en: {output_file}!")

except Exception as e:
    print(f"Error al generar el archivo Excel: {e}")
    print("Intentando generar archivo Excel básico...")
    try:
        # Generar archivo Excel básico sin formato
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in hojas.items():
                if df is not None and not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            if 'resumen_stock_final' in locals() and not resumen_stock_final.empty:
                resumen_stock_final.to_excel(writer, sheet_name="RESUMEN_STOCK", index=False)
        print(f"Archivo Excel básico generado en: {output_file}")
    except Exception as e2:
        print(f"Error al generar archivo Excel básico: {e2}")
        sys.exit(1)

# --- FIN: Generar Excel con formato personalizado ---
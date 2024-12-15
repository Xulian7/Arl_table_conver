#region librerias
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Label, Button, Frame, Entry
from tkinter.ttk import Progressbar
from rapidfuzz import fuzz, process
from fuzzywuzzy import process
import re
import json
import calendar
from datetime import datetime
import os
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image

#endregion

#region procesamiento xlsx

def process_file(file_path, file_path2, output_folder, keywords, threshold):
    try:
        # Ruta de salida
        output_path = f"{output_folder}/Reporte de inactividades.xlsx"

        def read_excel_file(file_path):
            #Función para leer una hoja de cálculo Excel con el motor correcto según la extensión del archivo.
            if file_path.endswith('.xls'):
                return pd.read_excel(file_path, engine='xlrd')
            else:
                return pd.read_excel(file_path, engine='openpyxl')

        # Ejemplo de uso con las rutas de archivos
        df = read_excel_file(file_path)
        df2 = read_excel_file(file_path2)

        # Estadísticas iniciales
        total_rows = len(df)
        initial_rows = total_rows



#region data frame de nomina

        # Función genérica para asignar fechas según condiciones
        def asignar_fecha(fila, columna_fecha, referencia, limite_superior=True):
        
            try:
                # Verificar si la columna tiene un valor nulo
                if pd.isnull(fila[columna_fecha]):
                    return referencia
                
                # Asignar fecha según el límite especificado
                if limite_superior:
                    return max(fila[columna_fecha], referencia)
                else:
                    return min(fila[columna_fecha], referencia)
            except Exception as e:
                print(f"Error al procesar la fila en {columna_fecha}: {e}")
                return None


        # Configuración de fechas de referencia
        fecha_1_enero = datetime(datetime.now().year, 1, 1)
        fecha_hoy = datetime.now()

        # Crear columnas nuevas aplicando la función genérica
        df2['Fecha1'] = df2.apply(asignar_fecha, axis=1, args=('FECHA INGRESO', fecha_1_enero, True))
        df2['Fecha2'] = df2.apply(asignar_fecha, axis=1, args=('FECHA RETIRO', fecha_hoy, False))
        df2['Fecha1'] = pd.to_datetime(df2['Fecha1'], errors='coerce')
        df2['Fecha2'] = pd.to_datetime(df2['Fecha2'], errors='coerce')

        meses = [calendar.month_name[i] for i in range(1, 13)]
        for mes in meses:
            df2[mes] = 0

        for index, row in df2.iterrows():
            start_date = row['Fecha1']
            end_date = row['Fecha2']
            if pd.notna(start_date) and pd.notna(end_date) and start_date <= end_date:
                # Iterar mes por mes dentro del rango
                current_date = start_date
                while current_date <= end_date:
                    # Obtener el nombre del mes actual
                    mes = calendar.month_name[current_date.month]
                    # Incrementar el conteo en la columna correspondiente al mes
                    df2.at[index, mes] += 1
                    # Avanzar al siguiente día
                    current_date += pd.Timedelta(days=1)

        #Calcular FTE x mes
        def calcular_porcentaje(df, meses):
            for mes in meses:
                # Obtener el número de días del mes correspondiente
                mes_numero = list(calendar.month_name).index(mes)
                anio_actual = datetime.now().year  # Puedes ajustar al año que corresponda
                dias_del_mes = calendar.monthrange(anio_actual, mes_numero)[1]
                   
                # Calcular el porcentaje y actualizar los valores
                df[mes] = df[mes] / dias_del_mes# Convertir a porcentaje como decimal
            return df

        # Llamar a la función
        df2 = calcular_porcentaje(df2, meses)

#endregion

        # Crear una columna concatenada de 'IDENTIFICACION' y 'REAL INICIO'
        # Eliminar duplicados basados en la columna concatenada
        df['concat'] = df['IDENTIFICACION'].astype(str) + df['REAL INICIO'].astype(str)
        duplicates_found = len(df) - len(df.drop_duplicates(subset=['concat'], keep='first'))
        df.drop_duplicates(subset=['concat'], inplace=True)
        df.drop(columns=['concat'], inplace=True)

        # Función para asignar categorías según palabras clave
        def categorize_reason(reason):
            if not reason or pd.isna(reason):
                return "delete"
            
            normalized_reason = reason.lower().strip()
            
            for keyword, category in keywords:
                normalized_keyword = keyword.lower().strip()
                result = process.extractOne(normalized_reason, [normalized_keyword], scorer=fuzz.partial_ratio)

                if result and result[1] > threshold:
                    return category
            return "delete"
        
        #Asignar categorías basadas en palabras clave y eliminar = delete
        df['Categoria'] = df['MOTIVO'].astype(str).apply(categorize_reason)
        filtered_rows = df[df['Categoria'] != 'delete']
        # Eliminar las filas donde la categoría sea "delete"
        df = df[df['Categoria'] != 'delete']

        # Rutina para tratamiento de rango de fechas
        # Verificar y convertir las columnas de fechas
        df['REAL INICIO'] = pd.to_datetime(df['REAL INICIO'], errors='coerce')
        df['REAL FINAL'] = pd.to_datetime(df['REAL FINAL'], errors='coerce')

        # Verificar si hay fechas inválidas
        invalid_start_dates = df['REAL INICIO'].isna().sum()
        invalid_end_dates = df['REAL FINAL'].isna().sum()

        # Mostrar solo la suma de fechas inválidas
        if invalid_start_dates > 0 or invalid_end_dates > 0:
            label_status.config(
                text=f"Fechas inválidas encontradas : {invalid_start_dates + invalid_end_dates}",fg="red")

        # Eliminar filas con fechas inválidas
        df.dropna(subset=['REAL INICIO', 'REAL FINAL'], inplace=True)
        # Crear columnas para cada mes del año con los nombres reales
        # meses = [calendar.month_name[i] for i in range(1, 13)]
        for mes in meses:
            df[mes] = 0

        # Calcular días en cada mes dentro del rango de fechas
        for index, row in df.iterrows():
            start_date = row['REAL INICIO']
            end_date = row['REAL FINAL']
            if pd.notna(start_date) and pd.notna(end_date) and start_date <= end_date:
                # Iterar mes por mes dentro del rango
                current_date = start_date
                while current_date <= end_date:
                    # Obtener el nombre del mes actual
                    mes = calendar.month_name[current_date.month]
                    # Incrementar el conteo en la columna correspondiente al mes
                    df.at[index, mes] += 1
                    # Avanzar al siguiente día
                    current_date += pd.Timedelta(days=1)
        
        

        #Calcula dias transcurridos 
        df['DIAS_TRANSCURRIDOS'] = (df['REAL FINAL'] - df['REAL INICIO']).dt.days + 1

        # Crear un diccionario para mapear código -> descripción desde el JSON
        codes = {str(item['code']): item['desc'] for item in icd10_dict}  # Diccionario con código como clave y descripción como valor

        # Función para buscar código y descripción
        def find_code_and_diagnosis_with_progress(motivo, index):
            for code, description in codes.items():
                if re.search(r'\b' + re.escape(code) + r'\b', motivo):  # Buscar la coincidencia exacta con el código
                    progress_bar['value'] = index + 1  # Actualizar progreso visual
                    root.update_idletasks()
                    return code, description
            progress_bar['value'] = index + 1  # Actualizar progreso visual incluso si no se encuentra el código
            root.update_idletasks()
            return None, None

        # Iterar sobre la columna 'MOTIVO' para asignar 'CODE' y 'Diagnostico' al DataFrame
        df[['CODE', 'Diagnostico']] = [find_code_and_diagnosis_with_progress(motivo, index) for index, motivo in enumerate(df['MOTIVO'])]

        # Lista de posibles sedes
        sedes = ["ADMINISTRATIVA", "TOCANCIPÁ", "ITAGUI", "MONTEVIDEO", "SIBERIA", "AUTOSUR-BOSA"]
        cat_add = ['N° DE ENFERMEDADES LABORALES (E.L.)', 'N° DÍAS INCAPACIDAD POR ENFERMEDAD LABORAL']

        # Función para asignar la sede en base a coincidencias aproximadas y reglas específicas
        def asignar_sede(nomina):
            # Verificar coincidencia directa con "AUTOSUR" o "BOSA"
            if re.search(r'AUTOSUR', nomina, re.IGNORECASE) or re.search(r'BOSA', nomina, re.IGNORECASE):
                return 'AUTOSUR - BOSA'
            # Buscar la coincidencia más cercana en la lista de sedes
            sede, score = process.extractOne(nomina, sedes)
            return sede if score >= 80 else None  # Solo asigna la sede si la similitud es alta (umbral: 80%)

        # Crear la nueva columna 'SEDE' con la lógica ajustada
        df['SEDE'] = df['NÓMINA'].apply(asignar_sede)

        # Combina los dataframes de ausentismos y salarios 
        df = df.merge(df2[['IDENTIFICACION', 'SUELDO', 'CARGO']], on='IDENTIFICACION', how='left')
        # Convertir las columnas 'REAL INICIO' y 'REAL FINAL' a tipo fecha
        df['REAL INICIO'] = pd.to_datetime(df['REAL INICIO'], errors='coerce').dt.strftime('%d/%m/%Y')
        df['REAL FINAL'] = pd.to_datetime(df['REAL FINAL'], errors='coerce').dt.strftime('%d/%m/%Y')
        df['Salario_base'] = df['SUELDO']/30


        df['CostBrut_AT'] = df['DIAS_TRANSCURRIDOS']*df['Salario_base']
        df['CostBrut_AC'] = df['DIAS_TRANSCURRIDOS']*df['Salario_base']*0.66

        # Columnas de meses
        meses = ['January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December']

        # Crear la nueva columna 'Primer_Mes'
        df['Primer_Mes'] = df[meses].apply(
            lambda row: next((mes for mes, val in zip(meses, row) if val != 0), None),
            axis=1
        )   

        # Definir el orden de las columnas que deseas mantener
        column_order = [
            'IDENTIFICACION', 'NOMBRE COMPLETO', 'SEDE', 'SUELDO', 'Salario_base', 'CARGO', 
            'REAL INICIO', 'REAL FINAL', 'CLASE', 
            'CODE', 'Diagnostico', 'VALOR', 'Categoria', 'January', 'February', 
            'March', 'April', 'May', 'June', 'July', 'August', 'September', 
            'October', 'November', 'December', 'DIAS_TRANSCURRIDOS', 'Primer_Mes', 'CostBrut_AT', 'CostBrut_AC'
        ]

        column_order2 = [
            'SEDE', 'Categoria', 'January', 'February', 
            'March', 'April', 'May', 'June', 'July', 'August', 'September', 
            'October', 'November', 'December'
        ]

        column_order3 = ['SEDE', 'Categoria', 'Primer_Mes']

        # Reordenar el DataFrame y guardarlo en df_f
        df_f = df[column_order]

        df_filtered = df_f.melt(id_vars=['SEDE', 'Categoria'], 
                      value_vars=['January', 'February', 'March', 'April', 'May', 'June', 
                                  'July', 'August', 'September', 'October', 'November', 'December'],
                      var_name='Mes', value_name='Valor')

        # Lista de categorías a eliminar
        categorias_a_eliminar = ['LIC. MAT', 'LIC. PAT']
        # Filtrar el DataFrame quitando las categorías no deseadas
        df_filtered = df_filtered[~df_filtered['Categoria'].isin(categorias_a_eliminar)]
        
        # Crear todas las combinaciones posibles de SEDE y Categoria
        unique_categorias = ["A.C.", "A.T."]

        all_combinations = pd.MultiIndex.from_product(
            [sedes, unique_categorias],
            names=["SEDE", "Categoria"]
        )
        
        # Crear DataFrame con todas las combinaciones
        base_df = pd.DataFrame(index=all_combinations)

        # 
        pivot_table = df_filtered.pivot_table(
            index=['SEDE', 'Categoria'],  # Filas: SEDE y Categoria
            columns='Mes',               # Columnas: Mes
            values='Valor',              # Valores: Valor
            aggfunc='sum',               # Agregación: suma
            fill_value=0                 # Llenar valores faltantes con 0
        )

        # Asegurarse de que todas las combinaciones estén presentes en la pivot table
        pivot_table = base_df.join(pivot_table, how="left").fillna(0).reset_index()
        # Convertir los valores de la pivot table a enteros
        pivot_table.iloc[:, 2:] = pivot_table.iloc[:, 2:].astype(int)
        #Reordenar meses
        pivot_table =pivot_table[column_order2]
        #Generar dataframe para tabla dinamica #2 x agregacion de conteo
        df_aux = df[column_order3]
        df_aux = df_aux[~df_aux['Categoria'].isin(categorias_a_eliminar)]

        months_order = ['January', 'February', 'March', 'April', 'May', 'June', 
        'July', 'August', 'September', 'October', 'November', 'December']

        # Crear la tabla dinámica
        pivot_table2 = pd.pivot_table(
            df_aux, 
            index=['SEDE', 'Categoria'],  # Filas
            columns='Primer_Mes',         # Columnas (los meses)
            aggfunc='size',               # Conteo
            fill_value=0                  # Llenar valores nulos con 0
        )

        # Reindexar las columnas para asegurarnos de incluir todos los meses
        pivot_table2 = pivot_table2.reindex(columns=months_order, fill_value=0)

        # Asegurarse de que todas las combinaciones estén presentes en la pivot table
        pivot_table2 = base_df.join(pivot_table2, how="left").fillna(0).reset_index()

        # Convertir los valores de la pivot table a enteros
        pivot_table2.iloc[:, 2:] = pivot_table2.iloc[:, 2:].astype(int)
    
        # Estadísticas finales
        final_rows = len(df)
        rows_filtered = initial_rows - final_rows
        stats = {
            "Total de registros iniciales": initial_rows,
            "Registros filtrados por palabras clave": initial_rows - len(filtered_rows),
            "Duplicados eliminados": duplicates_found,
            "Registros finales": final_rows
        }
        
        # Crear una nueva columna basada en condiciones TABLA DIAS TOTALES
        def categorize_s(value):
            if value == "A.C.":
                return "(A.C.) DIAS INACAPACIDAD ACCIDENTES COMUNES"
            elif value == "A.T.":
                return "(A.T.) DIAS INCAPACIDAD ACCIDENTES DE TRABAJO"
            else:
                return value  # Si no coincide, se deja igual

        pivot_table["Categoria"] = pivot_table["Categoria"].apply(categorize_s)
        # Crear una nueva columna basada en condiciones TABLA CANTIDAD INCAPACIDADES
        def categorize_n(value):
            if value == "A.C.":
                return "(A.C.) N° DE ACCIDENTES COMUNES"
            elif value == "A.T.":
                return "(A.T.) N° DE ACCIDENTES DE TRABAJO"
            else:
                return value  # Si no coincide, se deja igual

        pivot_table2["Categoria"] = pivot_table2["Categoria"].apply(categorize_n)

        # Crear un DataFrame vacío con los índices y las columnas
        data = []

        # Repetir las sedes para las dos categorías y agregar ceros a las columnas de meses
        for sede in sedes:
            for categoria in cat_add:
                row = [sede, categoria] + [0] * len(meses)  # Agregar ceros por cada mes
                data.append(row)

        # Crear el DataFrame
        df_add = pd.DataFrame(data, columns=['SEDE', 'Categoria'] + meses)
        df_report = pd.concat([pivot_table, pivot_table2, df_add], ignore_index=True)
        df_report = df_report.sort_values(by=['SEDE', 'Categoria'])
        
        # Agrupar por la columna 'Categoria' y sumar las columnas de meses
        df_sumado = df_report.groupby('Categoria').sum(numeric_only=True).reset_index()
        df_report2 = pd.concat([df_report, df_sumado], ignore_index=True)

        with pd.ExcelWriter(output_path) as writer:
            df_f.to_excel(writer, sheet_name='Datos', index=False)
            df2.to_excel(writer, sheet_name='Horas', index=False)
            df_report2.to_excel(writer, sheet_name='Report', index=False)

     

        return True, stats
    except Exception as e:
        print(e)
        return False, {"error": str(e)}   
#endregion

#region funcionalidad botones
def load_json():
    global icd10_dict
    try:
        with open("cie10.json", "r", encoding="utf-8") as json_file:
            icd10_dict = json.load(json_file)  # Aquí se cargará la lista
            # Mostrar los primeros elementos de la lista para validación
            label_json_display.config(text="Diccionario cie_10 cargado con exito")
    except Exception as e:
        label_json_display.config(text=f"Error al cargar el archivo JSON: {str(e)}", fg="red")

#Funciones para seleccionar y cargar los archivos a la fila de procesos
def select_file(label, global_var_name=None, load_json_on_select=False):
    #Función genérica para seleccionar un archivo Excel y actualizar la etiqueta correspondiente.
    #Args:
    #    label (tk.Label): Etiqueta a actualizar con la ruta del archivo.
    #    global_var_name (str): Nombre de la variable global donde almacenar la ruta del archivo seleccionado (opcional).
    #    load_json_on_select (bool): Indica si se debe cargar el JSON automáticamente tras seleccionar el archivo.
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    if file_path:
        label.config(text=file_path, fg="black")
        if global_var_name:
            globals()[global_var_name] = file_path  # Almacenar la ruta en la variable global correspondiente
        if load_json_on_select:
            load_json()
    else:
        label.config(text="SIN ARCHIVO", fg="red")

def select_file_ausentismos():
    #Función específica para seleccionar el archivo de ausentismos.
    #Llama a la función genérica y carga el JSON automáticamente después.
        select_file(label_file_path, global_var_name="selected_file_path", load_json_on_select=True)

def select_file_personal():
    #Función específica para seleccionar el archivo de personal.
    #Llama a la función genérica.
    select_file(label_file_path2, global_var_name="selected_file_path2")

def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory()
    if output_folder_path:
        label_output_folder.config(text=output_folder_path, fg="blue", cursor="hand2", font=("Helvetica", 10))
    else:
        label_output_folder.config(text="DETERMINAR RUTA DE SALIDA", fg="red")

# Abre la carpeta de destino en el explorador de archivos
def open_folder(event):
    folder_path = label_output_folder.cget("text")  # Obtener la ruta del texto del label
    if os.path.exists(folder_path):
        os.startfile(folder_path) 


def load_combine_and_save_files():
    #Selecciona múltiples archivos Excel, los carga, los combina y permite guardar el resultado.
    global label_status
    # Seleccionar múltiples archivos
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])

    if file_paths:
        try:
            # Mostrar estado inicial en la etiqueta
            label_status.config(text=f"{len(file_paths)} archivo(s) seleccionado(s). Procesando...", fg="black")
            label_status.update()  # Actualizar la interfaz

            # Cargar y combinar archivos
            dataframes = [pd.read_excel(path) for path in file_paths]
            combined_df = pd.concat(dataframes, ignore_index=True)

            # Abrir diálogo para guardar el archivo combinado
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if file_path:
                # Guardar el DataFrame combinado
                combined_df.to_excel(file_path, index=False)
                label_status.config(text=f"Archivo combinado guardado exitosamente en: {file_path}", fg="green")
            else:
                label_status.config(text="Guardado cancelado.", fg="orange")
        except Exception as e:
            # Mostrar error si algo falla
            label_status.config(text=f"Error al procesar archivos: {e}", fg="red")
    else:
        label_status.config(text="SIN ARCHIVOS SELECCIONADOS", fg="red")

#Funcionalidad del Boton que da inicio al analisi
def clean_file():
    if not selected_file_path:
        label_file_path.config(text="Debe seleccionar un archivo ausentismos", fg="red")
        return
    if not selected_file_path2:
        label_file_path2.config(text="Debe seleccionar un archivo personal", fg="red")
        return
    if not output_folder_path:
        label_output_folder.config(text="DETERMINAR RUTA DE SALIDA", fg="red")
        return

    try:
        threshold = int(80)
    except ValueError:
        label_status.config(text="El umbral debe ser un número entero", fg="red")
        return

    success, stats = process_file(selected_file_path, selected_file_path2, output_folder_path, keywords, threshold)
    if success:
        label_status.config(text="Procesamiento completado con éxito", fg="green")
        # Actualizar las estadísticas
        label_stats.config(
            text=f"Total registros: {stats['Total de registros iniciales']}\n"
                 f"Registros filtrados: {stats['Registros filtrados por palabras clave']}\n"
                 f"Duplicados eliminados: {stats['Duplicados eliminados']}\n"
                 f"Registros finales: {stats['Registros finales']}",
            fg="blue"
        )
    else:
        label_status.config(text=f"Error: {stats.get('error', 'Desconocido')}", fg="red")

def reset_labels():
    label_file_path.config(text="En espera de Archivo de ausentismos", fg="blue")
    label_file_path2.config(text="En espera de archivo de personal", fg="blue")
    label_output_folder.config(text="En espera seleccionar Ruta de Salida", fg="blue", underline=0)
    label_status.config(text="", fg="black")

#endregion

#region tkinter

# Crear la ventana principal



root = tk.Tk()
root.title("Procesador de Archivos Excel Ausentismos")
root.geometry("900x800")
root.iconbitmap("img/banner.ico")

# Variables globales
selected_file_path = None
selected_file_path2 = None
output_folder_path = None
keywords = [
    ("Accidente de trabajo", "A.T."),
    ("Incapacidad", "A.C."),
    ("Licencia paternidad", "LIC. PAT"),
    ("Licencia maternidad", "LIC. MAT")
]

# Construir rutas dinámicamente
current_dir = os.path.dirname(os.path.abspath(__file__))
filex_path = os.path.join(current_dir, "img", "ausentismos.png")
filey_path = os.path.join(current_dir, "img", "personal.png")
folderx_path = os.path.join(current_dir, "img", "folder.png")
datax_path = os.path.join(current_dir, "img", "data.png")
multi_path = os.path.join(current_dir, "img", "multi.png")
escoba_path =os.path.join(current_dir, 'img', "escoba.png")

# Cargar las imágenes con rutas absolutas
filex = tk.PhotoImage(file=filex_path)
filey = tk.PhotoImage(file=filey_path)
folderx = tk.PhotoImage(file=folderx_path)
datax = tk.PhotoImage(file=datax_path)
multi = tk.PhotoImage(file=multi_path)
escoba = tk.PhotoImage(file=escoba_path)

# Configurar pesos en el root para adaptabilidad
root.grid_rowconfigure(0, weight=1)  # Frame botones
root.grid_rowconfigure(1, weight=1)  # Frame etiquetas
root.grid_rowconfigure(2, weight=1)  # Frame inferior
root.grid_columnconfigure(0, weight=1)

# Crear Frames para organizar los widgets
frame_buttons = tk.Frame(root)
frame_buttons.grid(row=0, column=0, pady=5, padx=20, sticky="nsew")
frame_labels = tk.Frame(root, relief='groove', bd=2)
frame_labels.grid(row=1, column=0, pady=5, padx=20, sticky="nsew")
frame_bottom = tk.Frame(root, relief='groove', bd=1)
frame_bottom.grid(row=2, column=0, pady=5, padx=20, sticky="nsew")

# Configurar pesos en frame_buttons
frame_buttons.grid_rowconfigure(0, weight=1)  # Botones
frame_buttons.grid_rowconfigure(1, weight=1)  # Barra de progreso
for col in range(5):  # Una columna para cada botón
    frame_buttons.grid_columnconfigure(col, weight=1)

# Configurar pesos en frame_labels
for row in range(4):  # Etiquetas
    frame_labels.grid_rowconfigure(row, weight=1)
frame_labels.grid_columnconfigure(0, weight=1)

# Configurar pesos en frame_bottom
for row in range(4):  # Widgets de la parte inferior
    frame_bottom.grid_rowconfigure(row, weight=1)
frame_bottom.grid_columnconfigure(0, weight=1)

# Frame 1: Botones en disposición horizontal con grid
btn_select_file = tk.Button(frame_buttons, width=80, height=80, text="Ausentismos", image=filex, compound='bottom', command=select_file_ausentismos)
btn_select_file.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

btn_select_file2 = tk.Button(frame_buttons, width=80, height=80, text="Personal", image=filey, compound='bottom', command=select_file_personal)
btn_select_file2.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

btn_select_output = tk.Button(frame_buttons, width=80, height=80, text="Carpeta Salida", image=folderx, compound='bottom', command=select_output_folder)
btn_select_output.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")

btn_multi = tk.Button(frame_buttons, width=80, height=80, text="Agrupar Archivos", image=multi, compound='bottom', command=load_combine_and_save_files)
btn_multi.grid(row=0, column=3, padx=10, pady=10, sticky="nsew")

btn_clean_file = tk.Button(frame_buttons, width=80, height=80, text="Procesar", image=datax, compound='bottom', command=clean_file)
btn_clean_file.grid(row=0, column=4, padx=10, pady=10, sticky="nsew")

progress_bar = Progressbar(frame_buttons, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=1, column=0, columnspan=5, pady=10, sticky="ew")


# Frame 2: Labels

btn_erase = tk.Button(frame_labels, width=80, height=80, text="Reiniciar rutas", image=escoba, compound='bottom', command=reset_labels)
btn_erase.grid(row=0, column=0, padx=10, pady=10, sticky="n")
label_file_path = tk.Label(frame_labels, text="En espera de Archivo de ausentismos", fg="blue")
label_file_path.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
label_file_path2 = tk.Label(frame_labels, text="En espera de archivo de personal", fg="blue")
label_file_path2.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
label_output_folder = tk.Label(frame_labels, text="En espera seleccionar Ruta de Salida", fg="blue", cursor="hand2")
label_output_folder.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
label_output_folder.bind("<Button-1>", open_folder)
label_status = tk.Label(frame_labels, text="", fg="black")
label_status.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
# Cambiar colores en los widgets


# Frame 3: Etiquetas y otros widgets en la parte inferior
label_stats = tk.Label(frame_bottom, text="", justify="left", fg="blue")
label_stats.grid(row=0, column=0, pady=5, sticky="nsew")
label_json_display = tk.Label(frame_bottom, text="Diccionario ICD_10", justify="left", fg="blue")
label_json_display.grid(row=1, column=0, pady=5, sticky="nsew")

# Iniciar el loop principal
root.mainloop()

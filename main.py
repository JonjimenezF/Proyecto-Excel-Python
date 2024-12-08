import tkinter as tk
from tkinter import messagebox, Toplevel, ttk
import pandas as pd
from datetime import datetime
import os
import pyodbc
import re

def conectar_base_datos():
    """Establece la conexión con la base de datos Access."""
    try:
        ruta_db = os.path.join(os.path.dirname(__file__), 'estadistica compra', 'BASE.accdb')
        conn = pyodbc.connect(
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={ruta_db};"
        )
        return conn
    except Exception as e:
        raise Exception(f"No se pudo conectar a la base de datos: {e}")

def obtener_columnas(tabla):
    """Obtiene las columnas de la tabla en Access."""
    conn = conectar_base_datos()
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT TOP 1 * FROM {tabla}")
        columnas = [column[0] for column in cursor.description]
        conn.close()
        return columnas
    except Exception as e:
        raise Exception(f"Error al obtener columnas: {e}")

def obtener_datos(tabla):
    """Obtiene los datos completos de la tabla en Access."""
    conn = conectar_base_datos()
    try:
        query = f"SELECT * FROM {tabla}"
        df = pd.read_sql(query, conn)
        conn.close()

        # Asegúrate de convertir la columna 'numero_bodega' a tipo entero
        if 'numero_bodega' in df.columns:
            df['numero_bodega'] = df['numero_bodega'].astype('Int64')  # Usa 'Int64' para permitir valores nulos si es necesario

        return df
    except Exception as e:
        raise Exception(f"Error al obtener datos: {e}")

def aplicar_filtro(df, filtros):
    """Aplica múltiples filtros a un DataFrame y muestra los resultados en la interfaz."""
    try:
        df_filtrado = df
        for campo, valor in filtros:
            df_filtrado = df_filtrado[df_filtrado[campo].astype(str) == str(valor)]

        if df_filtrado.empty:
            messagebox.showinfo("Resultado", "No se encontraron registros para los filtros aplicados.")
        else:
            messagebox.showinfo("Éxito", "Datos filtrados aplicados correctamente.")
        
        return df_filtrado
    except Exception as e:
        messagebox.showerror("Error", f"Error al aplicar los filtros: {e}")
        return None


def agregar_filtro():
    """Agrega un filtro a la lista de filtros activos y actualiza la interfaz."""
    campo = combo_filtro_campo.get()
    valor = combo_filtro_valor.get()

    if not campo or not valor:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un campo y un valor para el filtro.")
        return

    filtros_activos.append((campo, valor))
    actualizar_lista_filtros()

def eliminar_filtro(filtro):
    filtro_tuple = tuple(filtro.split('='))
    filtro_tuple = (filtro_tuple[0].strip(), filtro_tuple[1].strip())
    
    if filtro_tuple in filtros_activos:
        filtros_activos.remove(filtro_tuple)
        
        # Actualizar la interfaz
        lista_filtros.delete(0, tk.END)  # Borra todos los elementos
        for item in filtros_activos:
            lista_filtros.insert(tk.END, f"{item[0]} = {item[1]}")
    else:
        print(f"Filtro '{filtro}' no encontrado en la lista.")


def actualizar_lista_filtros():
    """Actualiza la interfaz para mostrar los filtros activos."""
    lista_filtros.delete(0, tk.END)
    for filtro in filtros_activos:
        lista_filtros.insert(tk.END, f"{filtro[0]} = {filtro[1]}")

def calcular_ultimos_12_meses(row, df):
    columnas_meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic"]
    ano_actual = datetime.now().year
    mes_actual = datetime.now().month
    valores_ultimos_12 = [0] * 12

    # Obtener índice de la fila actual
    fila_actual_index = row.name

    # Procesar el año actual
    if row["Ano"] == ano_actual:
        for i, col in enumerate(columnas_meses[:mes_actual]):
            if col in row.index and not pd.isnull(row[col]):
                # Calcular el índice de la lista de manera segura
                index = mes_actual - 1 - i
                if 0 <= index < 12:
                    valores_ultimos_12[index] = int(row[col])

    # Buscar la fila del año anterior en la misma posición
    if row["Ano"] == ano_actual:
        try:
            # Buscar la fila anterior en el DataFrame usando el índice correcto
            fila_anterior = df[df.index == fila_actual_index - 1]  # Filtra por el índice correcto
            if not fila_anterior.empty:
                fila_anterior = fila_anterior.iloc[0]  # Obtén la primera fila de la búsqueda
                if fila_anterior["Código"] == row["Código"] and fila_anterior["Ano"] == ano_actual - 1:
                    meses_restantes = 12 - mes_actual
                    for i, col in enumerate(columnas_meses[-meses_restantes:]):
                        if col in fila_anterior.index and not pd.isnull(fila_anterior[col]):
                            # Calcular el índice de la lista de manera segura
                            index = mes_actual + i
                            if 0 <= index < 12:
                                valores_ultimos_12[index] = int(fila_anterior[col])
        except KeyError:
            # Maneja el caso en que no haya una fila anterior válida
            pass

    # Retornar la suma total de los últimos 12 meses
    return sum(valores_ultimos_12)
def cargar_stock_desde_excel(ruta_excel):
    """
    Carga los datos de stock desde un archivo Excel y asigna los valores de stock a las correspondientes bodegas.
    Elimina las filas duplicadas basadas en el código de producto, manteniendo solo la información.
    """
    try:
        # Cargar el archivo Excel y la hoja especificada
        df_stock = pd.read_excel(ruta_excel, sheet_name="Reformateado")

        # Imprimir los nombres de las columnas para depurar
        #print("Nombres de columnas en el archivo:", df_stock.columns)

        # Verificar que la columna 'CódigoProducto' exista y renombrar
        if "CódigoProducto" not in df_stock.columns:
            raise Exception("El archivo Excel debe contener la columna 'CódigoProducto'.")
        
        # Verificar que la columna 'Bodega' exista
        if "Bodega" not in df_stock.columns:
            raise Exception("El archivo Excel debe contener la columna 'Bodega'.")
        
        df_stock.rename(columns={"CódigoProducto": "Código"}, inplace=True)

        # Verificar que la columna 'Stock' exista
        if "Cantidad" not in df_stock.columns:
            raise Exception("El archivo Excel debe contener la columna 'Stock'.")

        # Eliminar filas duplicadas por el código de producto, manteniendo la primera ocurrencia
        df_stock = df_stock.drop_duplicates(subset=["Código", "Bodega"], keep="first")

        # Imprimir para verificar el resultado antes de regresar el DataFrame
        #print("Datos actualizados en el DataFrame con stock sin duplicados:", df_stock)

        # Retorna el DataFrame procesado
        return df_stock

    except Exception as e:
        raise Exception(f"Error al cargar y procesar los datos de stock desde Excel: {e}")



    

filtros_activos = []

def generar_reporte():
    global filtros_activos 
    # Campos obligatorios para los cálculos internos
    campos_obligatorios = ["Ano", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                           "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Código", "numero_bodega"]

    # Obtener los campos seleccionados desde la interfaz
    seleccionados = [campo for campo, var in campos_seleccionados.items() if var.get() == 1]

    if not seleccionados:
        messagebox.showwarning("Advertencia", "Selecciona al menos un campo.")
        return

    # Asegurarse de que los campos obligatorios siempre se añadan a la selección
    seleccionados = list(set(seleccionados + campos_obligatorios))

    # Obtener los campos de agrupación seleccionados
    campos_agrupacion_seleccionados = [campo for campo, var in campos_agrupacion.items() if var.get() == 1]
    if not campos_agrupacion_seleccionados:
        messagebox.showwarning("Advertencia", "Selecciona al menos un campo para agrupar.")
        return

    try:
        # Cargar los datos de la tabla de la base de datos
        tabla = "Estadistica"
        df = obtener_datos(tabla)
        # Aplicar filtro si se seleccionó un campo de filtro y un valor

        # Obtener múltiples filtros activos y aplicarlos
        filtros_activos = [(campo, valor) for campo, valor in filtros_activos]

        if filtros_activos:
            df = aplicar_filtro(df, filtros_activos)
            if df is None:  # Si los filtros no devuelven datos, salir de la función
                return

        
        # Filtrar columnas seleccionadas
        columnas_seleccionadas = [col for col in df.columns if col in seleccionados]
        df_calculos = df[columnas_seleccionadas]

        # Calcular los 12 meses
        df_calculos['12 Meses'] = df_calculos.apply(lambda row: calcular_ultimos_12_meses(row, df), axis=1)
        # Calcular los Total
        df_calculos["Total"] = df_calculos[["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic"]].sum(axis=1)

        # Cargar stock desde Excel y hacer la fusión
        ruta_stock = os.path.join(os.path.dirname(__file__), 'estadistica compra', 'Tabla dinámica Analysis (x_bi_sql_view.stock_bodega_sc).xlsm')
        df_stock = cargar_stock_desde_excel(ruta_stock)
        ano_actual = datetime.now().year
        df_stock['Ano'] = ano_actual
        df_stock.rename(columns={"Cantidad": "Stock"}, inplace=True)
        df_calculos = df_calculos.merge(df_stock, left_on=["Código", "numero_bodega", "Ano"], right_on=["Código", "Bodega", "Ano"], how="left")
        
        df_calculos['Prom'] = (df_calculos['12 Meses'] / 12).astype(float).round(0)
        #df_calculos['Dur'] = df_calculos['Stock'] / df_calculos['Prom'].replace(0)
        # Reemplazar valores de Prom = 0 por NaN para evitar divisiones por cero
        #df_calculos['Dur'] = (df_calculos['Stock'] / df_calculos['Prom'].replace(0, pd.NA)).fillna(0).replace([float('inf'), -float('inf')], 0).round(2)

         #df_calculos['Dur'] = df_calculos['Dur'].fillna(0).round(2)   
        # print(df_calculos[['Stock', 'Prom']], "     ", df_calculos['Stock'] / df_calculos['Prom'])
        # print (df_calculos['Dur'] )
        #Agrupar
        df_grouped = df_calculos.groupby(campos_agrupacion_seleccionados + ["Ano"]).agg({
            "Ene": "sum", 
            "Feb": "sum", 
            "Mar": "sum", 
            "Abr": "sum", 
            "May": "sum", 
            "Jun": "sum", 
            "Jul": "sum", 
            "Ago": "sum", 
            "Sept": "sum", 
            "Oct": "sum", 
            "Nov": "sum", 
            "Dic": "sum",
            "Total": "sum",
            "12 Meses": "sum",
            "Stock": "sum",
            "Prom": "sum"
        }).reset_index()

        #Calculo con los valores agrupados
        df_grouped['Dur'] = (
        df_grouped['Stock'] / df_grouped['Prom'].replace(0, pd.NA)
        ).fillna(0).replace([float('inf'), -float('inf')], 0).round(2)


        if agregar_subtotales_var.get() and agregar_subtotales_Ext.get():
            df_grouped = Combinar_funciones(df_grouped, campos_agrupacion_seleccionados)
            # Guardar el DataFrame resultante en un archivo Excel
            ruta_excel = "reporte_calculado_con_subtotales_y_extra.xlsx"
            df_grouped.to_excel(ruta_excel, index=False)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente con subtotales y extra: {ruta_excel}")

        elif agregar_subtotales_var.get():
            # Llamar a la función para calcular solo los subtotales
            df_grouped = agregar_subtotales(df_grouped, campos_agrupacion_seleccionados)
            
            # Guardar el DataFrame resultante en un archivo Excel
            ruta_excel = "reporte_calculado_con_subtotales.xlsx"
            df_grouped.to_excel(ruta_excel, index=False)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente con subtotales: {ruta_excel}")

        else:
            # Si no se marca el checkbox, guarda el reporte sin subtotales
            ruta_excel = "reporte_calculado.xlsx"
            df_grouped.to_excel(ruta_excel, index=False)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente: {ruta_excel}")

    except Exception as e:
        print("Error", f"Hubo un problema al generar el informe: {e}")
        messagebox.showerror("Error", f"Hubo un problema al generar el informe: {e}")



def agregar_subtotales(df, campos_agrupacion_seleccionados):
    """Genera subtotales por grupo de MEDID similares (por ejemplo, CLÁSICA) y agrega subtotales por año."""

    # Crear un DataFrame vacío para almacenar los resultados finales
    df_final = pd.DataFrame()

    # Agrupar por los campos de agrupación seleccionados y Año para obtener las sumas por cada combinación
    df_grouped = df.groupby(campos_agrupacion_seleccionados + ['Ano'], as_index=False).sum(numeric_only=True)

    # Función para obtener el nombre común del producto
    def obtener_nombre_comun(producto):
        # Aquí usamos una expresión regular para extraer solo la parte general del nombre (sin el tamaño)
        match = re.match(r"([A-Za-zÁÉÍÓÚáéíóú0-9]+(?: [A-Za-zÁÉÍÓÚáéíóú0-9]+)*)(?=\s*\d{3}X\d{3})", producto)
        if match:
            return match.group(1)  # Retorna el nombre común (por ejemplo, 'CLÁSICA')
        else:
            return producto  # Si no hay coincidencia, devuelve el nombre completo

    # Crear una nueva columna para identificar el grupo general del producto (por ejemplo, CLÁSICA, etc.)
    df_grouped['Grupo'] = df_grouped['MEDID'].apply(obtener_nombre_comun)

    # Iterar sobre los grupos generales de productos (por ejemplo, CLÁSICA, etc.)

    for grupo, data_grupo in df_grouped.groupby('Grupo'):
        # Agregar las filas originales para cada grupo
        for _, row in data_grupo.iterrows():
            df_final = pd.concat([df_final, row.to_frame().T], ignore_index=True)
        
        # Agregar subtotales por año para el grupo
        for ano in data_grupo['Ano'].unique():
            subtotales_ano = data_grupo[data_grupo['Ano'] == ano].sum(numeric_only=True)
            subtotales_ano['MEDID'] = f'Subtotal'  # Agregar la etiqueta de subtotal
            subtotales_ano['Ano'] = ano  # Mantener el año en los subtotales
            df_final = pd.concat([df_final, subtotales_ano.to_frame().T], ignore_index=True)
        

        # Agregar total por grupo (por ejemplo, "Total CLÁSICA")
        total_grupo = data_grupo.sum(numeric_only=True)
        total_grupo['MEDID'] = f'Total'  # Agregar la etiqueta de total
        total_grupo['Ano'] = ''  # Dejar vacío el año para los totales
        df_final = pd.concat([df_final, total_grupo.to_frame().T], ignore_index=True)
        
        # Agregar una fila vacía después del total
        fila_vacia = total_grupo.copy()
        fila_vacia[:] = ''  # Vaciar todos los valores de la fila
        df_final = pd.concat([df_final, fila_vacia.to_frame().T], ignore_index=True)

    # Eliminar la columna 'Grupo' antes de devolver el DataFrame
    df_final = df_final.drop(columns=['Grupo'], errors='ignore')  # Usamos 'errors=ignore' para evitar errores si no existe la columna

    return df_final


def Combinar_funciones(df, campos_agrupacion_seleccionados):
    """Genera subtotales por grupo de MEDID similares (por ejemplo, CLÁSICA) y agrega subtotales por año."""
    
    # Crear una lista para almacenar todos los resultados
    resultados = []

    # Agrupar por los campos de agrupación seleccionados y Año para obtener las sumas por cada combinación
    df_grouped = df.groupby(campos_agrupacion_seleccionados + ['Ano'], as_index=False).sum(numeric_only=True)

    # Función para obtener el nombre común del producto
    def obtener_nombre_comun(producto):
        match = re.match(r"([A-Za-zÁÉÍÓÚáéíóú0-9]+(?: [A-Za-zÁÉÍÓÚáéíóú0-9]+)*)(?=\s*\d{3}X\d{3})", producto)
        return match.group(1) if match else producto

    # Crear una nueva columna para identificar el grupo general del producto (por ejemplo, CLÁSICA, etc.)
    df_grouped['Grupo'] = df_grouped['MEDID'].apply(obtener_nombre_comun)

    # Iterar sobre los grupos generales de productos
    for grupo, data_grupo in df_grouped.groupby('Grupo'):
        # Lista temporal para almacenar las filas del grupo actual
        filas_grupo_temporal = []

        # Añadir las filas originales para cada grupo
        filas_grupo_temporal.extend(data_grupo.to_dict(orient='records'))

        # Agregar subtotales por año para el grupo (una sola vez por año)
        for ano in data_grupo['Ano'].unique():
            subtotales_ano = data_grupo[data_grupo['Ano'] == ano].sum(numeric_only=True)
            subtotales_ano['MEDID'] = f'TOTALES SUB'  # Agregar la etiqueta de subtotal con el nombre del grupo
            subtotales_ano['Ano'] = ano  # Mantener el año en los subtotales
            filas_grupo_temporal.append(subtotales_ano.to_dict())

        # Agregar total por grupo (por ejemplo, "Total CLÁSICA")
        total_grupo = data_grupo.sum(numeric_only=True)
        total_grupo['MEDID'] = f'TOTAL'  # Agregar la etiqueta de total con el nombre del grupo
        total_grupo['Ano'] = ''  # Dejar vacío el año para los totales
        filas_grupo_temporal.append(total_grupo.to_dict())

        # Función para extraer las medidas del nombre del producto
        def extraer_medidas(nombre_producto):
            match = re.search(r"(\d+)[Xx](\d+)", nombre_producto)
            if match:
                medida1 = int(match.group(1)) / 100  # Convertir a metros
                medida2 = int(match.group(2)) / 100  # Convertir a metros
                return medida1, medida2
            return None, None

        # Conjunto para realizar un seguimiento de las combinaciones únicas de 'MEDID' y 'Ano'
        combinaciones_vistas = set()
        print(data_grupo)

        # Ahora agregar las filas "Extra" (agrupando por Grupo)
        for ano in data_grupo['Ano'].unique():
            # Inicializar los valores de los meses a 0
            valores_totales_ano = {mes: 0 for mes in ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic","Total","12 Meses","Stock","Prom","Dur"]}
            total_meses_ano = 0
            meses_validos_ano = 0
            
            # Iterar sobre las filas para este año específico
            for _, row in data_grupo[data_grupo['Ano'] == ano].iterrows():
                # Extraer medidas y calcular "Extra"
                medida1, medida2 = extraer_medidas(row['MEDID'])
                if medida1 is not None and medida2 is not None:
                    # Calculando los valores de los meses para agrupar
                    for mes in valores_totales_ano:
                        cantidad_mes = pd.to_numeric(row[mes], errors='coerce')
                        if pd.notna(cantidad_mes) and cantidad_mes > 0:
                            valor_mes = round(medida1 * medida2 * cantidad_mes)
                            valores_totales_ano[mes] += valor_mes
                            total_meses_ano += valor_mes
                            meses_validos_ano += 1

            # Corregir los cálculos
            print(f"Calculando para el año {ano}:")
            print(f"Total de meses del año: {total_meses_ano}")
            print(f"Meses válidos del año: {meses_validos_ano}")
            fila_extra = {
                'MEDID': f"TOTALES EXTRA",  # Sin el grupo
                'Ano': ano,
                # 'Total': total_meses_ano,  # Total de todos los meses (ya calculado)
                # '12 Meses': total_meses_ano if meses_validos_ano > 0 else 0,  # Si hay meses válidos, usar el total, sino 0
                # 'Stock': round(total_meses_ano / 12) if meses_validos_ano > 0 else 0,  # Si hay meses válidos, calcular el stock
                # 'Prom': round(total_meses_ano / meses_validos_ano, 2) if meses_validos_ano > 0 else 0,  # Promedio por mes si hay meses válidos
                # 'Dur': round(total_meses_ano / 12, 2) if total_meses_ano > 0 else 0,  # Duración calculada si el total no es 0
            }
            # Agregar los valores de cada mes al diccionario de la fila
            for mes in valores_totales_ano:
                if valores_totales_ano[mes] == 0:
                    fila_extra[mes] = 0 
            fila_extra.update(valores_totales_ano)

            # Solo agregar la fila si no existe
            if (ano, grupo) not in combinaciones_vistas:
                filas_grupo_temporal.append(fila_extra)
                combinaciones_vistas.add((ano, grupo))
        # Agregar una fila vacía completa después de la fila "Extra"
        filas_grupo_temporal.append({columna: None for columna in filas_grupo_temporal[0].keys()})

        # Al finalizar el procesamiento de un grupo, agregar las filas del grupo a los resultados finales
        resultados.extend(filas_grupo_temporal)

    # Devolver el DataFrame final sin la columna 'Grupo'
    df_final = pd.DataFrame(resultados)
    df_final = df_final.drop(columns=['Grupo'], errors='ignore') 

    return df_final




# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Generar Reporte")
ventana.geometry("700x600")
# Obtener las dimensiones de la pantalla
pantalla_ancho = ventana.winfo_screenwidth()
pantalla_alto = ventana.winfo_screenheight()

# Obtener las dimensiones de la ventana
ventana_ancho = 700
ventana_alto = 600

# Calcular las coordenadas para centrar la ventana
pos_x = (pantalla_ancho // 2) - (ventana_ancho // 2)
pos_y = (pantalla_alto // 2) - (ventana_alto // 2)

ventana.geometry(f"{ventana_ancho}x{ventana_alto}+{pos_x}+{pos_y}")
ventana.configure(bg="#DFDDD9")  # Fondo de la ventana



# Obteniendo los campos de la tabla de la base de datos
campos = []
try:
    campos = obtener_columnas("Estadistica")
except Exception as e:
    messagebox.showerror("Error", f"No se pudieron obtener los campos: {e}")

campos_obligatorios = ["Ano", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                       "Jul", "Ago", "Sept", "Oct", "Nov", "Dic"]

campos_filtrados = [campo for campo in campos if campo not in campos_obligatorios]
campos_seleccionados = {}
campos_agrupacion = {}

# Estilo de ttk
style = ttk.Style()
style.theme_use('clam')
style.configure('TButton',
               font=('Segoe UI', 12),
               padding=6,
               relief="flat",
               background="#2196F3",  
               foreground="white")  
style.map('TButton',
          background=[('active', '#0b7dda')], 
          foreground=[('active', 'white')]) 

style.configure('TCombobox',
               padding=5,
               font=('Segoe UI', 10),
               background="#f0f4f8",
               borderwidth=1,
               relief="solid")
style.map('TCombobox',
          background=[('active', '#d9e1e8')], 
          foreground=[('active', 'black')])  

style.configure('TLabel',
               font=('Segoe UI', 12),
               background="#f0f4f8",
               foreground="black")


# Frame para la selección de campos con estilo moderno
frame_seleccion = ttk.LabelFrame(ventana, text="Selecciona los campos", padding=(10, 5), style="TFrame")
frame_seleccion.grid(row=0, column=0, rowspan=2, sticky="n", padx=10, pady=10)

# Crear y mostrar los checkbuttons de selección de campos con un estilo moderno
for i, campo in enumerate(campos_filtrados):
    var = tk.IntVar()
    checkbutton = ttk.Checkbutton(frame_seleccion, text=campo, variable=var, style="TCheckbutton")
    checkbutton.pack(side="top", anchor="w", padx=2, pady=2)
    campos_seleccionados[campo] = var

# Frame para los filtros
frame_filtro = ttk.LabelFrame(ventana, text="Filtrar por", padding=(10, 5), style="TFrame")
frame_filtro.grid(row=0, column=2, sticky="n", padx=10, pady=10)

# Dropdown para seleccionar el campo a filtrar
combo_filtro_campo = ttk.Combobox(frame_filtro, values=campos_filtrados)
combo_filtro_campo.pack(pady=5, fill="x")

# Dropdown para seleccionar un valor de la columna
combo_filtro_valor = ttk.Combobox(frame_filtro)
combo_filtro_valor.pack(pady=5, fill="x")

# Botón para agregar un filtro
btn_agregar_filtro = ttk.Button(frame_filtro, text="Agregar Filtro", command=agregar_filtro)
btn_agregar_filtro.pack(pady=5)

# Lista de filtros activos
lista_filtros = tk.Listbox(frame_filtro, font=("Segoe UI", 10), height=10, width=30)
lista_filtros.pack(pady=5)

# Botón para borrar el filtro seleccionado
btn_eliminar_filtro = ttk.Button(frame_filtro, text="Eliminar Filtro", command=lambda: eliminar_filtro(lista_filtros.get(tk.ACTIVE)))
btn_eliminar_filtro.pack(pady=5)


# Función para actualizar los valores del dropdown de valores
def actualizar_valores_dropdown(event):
    """Actualiza los valores en el dropdown de valores cuando se selecciona un campo."""
    campo = combo_filtro_campo.get()
    if campo and campo in campos_filtrados:
        df = obtener_datos("Estadistica")
        valores = df[campo].dropna().unique()
        valores = [str(valor) for valor in valores if not pd.isnull(valor)]
        combo_filtro_valor['values'] = valores

combo_filtro_campo.bind("<<ComboboxSelected>>", actualizar_valores_dropdown)

# Frame para la agrupación de campos con estilo
frame_agrupacion = ttk.LabelFrame(ventana, text="Agrupar por", padding=(10, 5), style="TFrame")
frame_agrupacion.grid(row=0, column=1, sticky="n", padx=10, pady=10)

# Crear y mostrar los checkbuttons de agrupación con un estilo más moderno
for i, campo in enumerate(campos_filtrados):
    var = tk.IntVar()
    checkbutton = ttk.Checkbutton(frame_agrupacion, text=campo, variable=var, style="TCheckbutton")
    checkbutton.pack(side="top", anchor="w", padx=2, pady=2)
    campos_agrupacion[campo] = var


agregar_subtotales_Ext = tk.IntVar()

check_subtotales_extra = tk.Checkbutton(ventana, text="Incluir subtotales extra", variable=agregar_subtotales_Ext)
check_subtotales_extra.grid(row=3, column=0, padx=1, pady=1)


agregar_subtotales_var = tk.IntVar()

check_subtotales = tk.Checkbutton(ventana, text="Incluir subtotales", variable=agregar_subtotales_var)
check_subtotales.grid(row=2, column=0, padx=10, pady=10)


ventana.grid_columnconfigure(0, weight=1)  
ventana.grid_columnconfigure(1, weight=1)  
ventana.grid_columnconfigure(2, weight=1)  

boton_agrupar = tk.Button(ventana, text="Generar informe", command=generar_reporte)
boton_agrupar.grid(row=4, column=1, pady=5, sticky="ew")


# Ejecutar la aplicación
ventana.mainloop()



import tkinter as tk
from tkinter import messagebox, Toplevel, ttk
import pandas as pd
from datetime import datetime
import os
import pyodbc
import re
import numpy as np

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

        #convertir la columna 'numero_bodega' a tipo entero
        if 'numero_bodega' in df.columns:
            df['numero_bodega'] = df['numero_bodega'].astype('Int64')  # Usa 'Int64' para permitir valores nulos si es necesario

        return df
    except Exception as e:
        raise Exception(f"Error al obtener datos: {e}")

def aplicar_filtro(df, filtros):
    """
    Aplica múltiples filtros a un DataFrame y muestra los resultados en la interfaz.
    Los filtros pueden incluir múltiples valores por columna.
    """
    try:
        # Crear una copia del DataFrame original para filtrar
        df_filtrado = df.copy()

        # Verificar el formato de los filtros
        if isinstance(filtros, list):
            # Convertir lista de tuplas en un diccionario {columna: [valores]}
            filtros_dict = {}
            for campo, valor in filtros:
                filtros_dict.setdefault(campo, []).append(valor)
        elif isinstance(filtros, dict):
            filtros_dict = filtros
        else:
            raise ValueError("El formato de 'filtros' debe ser una lista de tuplas o un diccionario.")

        # Aplicar filtros de forma acumulativa
        condicion_general = pd.Series(True, index=df.index)  # Iniciar con todas las filas como válidas
        for campo, valores in filtros_dict.items():
            # Convertir valores en una lista si no lo son
            if not isinstance(valores, list):
                valores = [valores]

            # Crear la condición para la columna actual
            condicion_columna = df[campo].astype(str).isin([str(valor) for valor in valores])

            # Combinar con las condiciones generales
            condicion_general &= condicion_columna

        # Filtrar el DataFrame usando las condiciones acumuladas
        df_filtrado = df[condicion_general]

        # Mostrar mensajes informativos
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


#-------------------------------------------------------------------------------------------------------------------------
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
    El stock solo se coloca una vez por codigo y N° bodega.
    """
    try:
        # Cargar el archivo Excel y la hoja especificada
        df_stock = pd.read_excel(ruta_excel, sheet_name="Reformateado")
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
                           "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Código", "numero_bodega","MT2"]

    # Antes de agrupar, incluye el campo seleccionado en los campos obligatorios
    campo_seleccionado = campos_agrupacion_selec.get()  # Obtén el campo seleccionado

    # Verifica si el campo seleccionado no está en los campos obligatorios
    if campo_seleccionado and campo_seleccionado not in campos_obligatorios:
        campos_obligatorios.append(campo_seleccionado)  # Añade el campo seleccionado

    # Ahora, la lista de campos obligatorios contendrá también el campo seleccionado
    campos_obligatorios = list(set(campos_obligatorios))  #     Elimina duplicados

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
        df_calculos = df_calculos.sort_values(by=["Código", "numero_bodega", "Ano"])
        df_calculos['Stock'] = df_calculos.groupby(['Código', 'numero_bodega', 'Ano'])['Stock'].transform(lambda x: x.where(x.index == x.index[0], 0))

        # Verificar el resultado
       

        mes_actual = datetime.now().month
        df_calculos['Prom'] = np.where(
            mes_actual != 0,  # Verifica que el mes no sea 0 (división por cero)
            round(df_calculos['12 Meses'] / mes_actual,3),  # Realiza la división
            np.nan  # Si mes_actual es 0, asigna NaN o un valor predeterminado
        )
        df_calculos['Prom'] = df_calculos['Prom'].replace([np.inf, -np.inf], np.nan)
        df_calculos['Prom'] = df_calculos['Prom'].fillna(0) 
        
        campo_seleccionado = campos_agrupacion_selec.get()  # Obtén el campo seleccionado
        if campo_seleccionado and campo_seleccionado not in campos_agrupacion_seleccionados:
            campos_agrupacion_seleccionados.append(campo_seleccionado)  # Asegúrate de que esté incluido

        df_grouped = df_calculos.groupby(campos_agrupacion_seleccionados + ["MT2"] + ["Ano"]).agg({
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
        df_grouped['Dur'] = np.where(
            df_grouped['Prom'] != 0,  # Asegúrate de que 'Prom' no sea 0
            round(df_grouped['Stock'] / df_grouped['Prom'],2),  # Realiza la división
            np.nan  # Si 'Prom' es 0, asigna NaN o un valor predeterminado
        )
        df_grouped['Dur'] = (df_grouped['Dur'])
        df_grouped = df_grouped.replace([np.inf, -np.inf], np.nan)  # Reemplaza infinitos por NaN
        df_grouped = df_grouped.fillna(0)
        #df_grouped['Dur'] = (
        #df_grouped['Stock'] / df_grouped['Prom'])
        # Formatear los valores numéricos con separadores de miles
       
        

        if agregar_subtotales_var.get() and agregar_subtotales_Ext.get():
            # Generar subtotales y guardar en un DataFrameW
            df_grouped = Combinar_funciones(df_grouped, campos_agrupacion_seleccionados)
            ruta_excel = "reporte_calculado_con_subtotales_y_extra.xlsx"
            # Aplicar formato a todas las columnas numéricas
            df_grouped.to_excel(ruta_excel, index=False)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente con totales y extra: {ruta_excel}")
        
        elif agregar_subtotales_Ext.get():
            
            
            #print (agregar_subtotales_Ext.get())
            # Guardar el DataFrame resultante en un archivo Excel
            
            df_grouped = agregar_subtotales_Extra(df_grouped, campos_agrupacion_seleccionados)

            ruta_excel = "reporte_calculado_con_extra.xlsx"
            df_grouped.to_excel(ruta_excel, index=False)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente con totales extra: {ruta_excel}")

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
            df_grouped = df_grouped.drop(columns=['MT2'], errors='ignore')
            df_grouped.to_excel(ruta_excel, index=False)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente: {ruta_excel}")
            # df_calculos = df_calculos.drop(columns=['MT2'], errors='ignore')
            # df_calculos.to_excel(ruta_excel, index=False)
            # messagebox.showinfo("Éxito", f"Reporte generado correctamente: {ruta_excel}")

    except Exception as e:
        print("Error", f"Hubo un problema al generar el informe: {e}")
        messagebox.showerror("Error", f"Hubo un problema al generar el informe: {e}")





def agregar_subtotales(df, campos_agrupacion_seleccionados):  # FUNCIÓN CON EL CRITERIO DE AGRUPACIÓN
    """Genera subtotales por grupo de MEDID similares (por ejemplo, CLÁSICA) y agrega subtotales por año."""

    # Crear un DataFrame vacío para almacenar los resultados finales
    df_final = pd.DataFrame()

    anos_presentes = df['Ano'].unique()

    # Agrupar por los campos de agrupación seleccionados y Año para obtener las sumas por cada combinación
    df_grouped = df.groupby(campos_agrupacion_seleccionados + ['Ano'], as_index=False).sum(numeric_only=True)

    # Función para obtener el nombre común del producto
    def obtener_nombre_comun(producto):
        # Usar expresión regular para extraer solo la parte general del nombre (sin el tamaño)
        match = re.match(r"([A-Za-zÁÉÍÓÚáéíóú0-9]+(?: [A-Za-zÁÉÍÓÚáéíóú0-9]+)*)(?=\s*\d{3}X\d{3})", producto)
        return match.group(1) if match else producto  # Retorna el nombre común o el nombre completo

    # Crear una nueva columna para identificar el grupo general del producto
    campo_agrupacion = campos_agrupacion_selec.get()  # Asumiendo que el usuario seleccionó solo un campo
    if campo_agrupacion == 'MT2':  
        df_grouped['Grupo'] = df_grouped['MT2']
    elif campo_agrupacion == "SUBFA":
        df_grouped['Grupo'] = df_grouped['SUBFA']
    elif campo_agrupacion == "LINEA":
        df_grouped['Grupo'] = df_grouped['LINEA']
    elif campo_agrupacion == "EMPRESA":
        df_grouped['Grupo'] = df_grouped['EMPRESA']
    elif campo_agrupacion == "COMP":
        df_grouped['Grupo'] = df_grouped['COMP']
    elif campo_agrupacion == "numero_bodega":
        df_grouped['Grupo'] = df_grouped['numero_bodega']
    elif campo_agrupacion == "FLIA":
        df_grouped['Grupo'] = df_grouped['FLIA']
    elif campo_agrupacion == "CATEGORIA":
        df_grouped['Grupo'] = df_grouped['CATEGORIA']
    elif campo_agrupacion == "PROVEEDOR":
        df_grouped['Grupo'] = df_grouped['PROVEEDOR']
    elif campo_agrupacion == "LOCAL":
        df_grouped['Grupo'] = df_grouped['LOCAL']
    elif campo_agrupacion == "MEDID":
        df_grouped['Grupo'] = df_grouped['MEDID']
    elif campo_agrupacion == "EMPRESA":
        df_grouped['Grupo'] = df_grouped['EMPRESA']
    else:
        if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
            df_grouped['Grupo'] = df_grouped['CRITERIO PARA AGRUPAR'].apply(obtener_nombre_comun)
        else:
            df_grouped['Grupo'] = df_grouped['DESCRIPCION']

    # Iterar sobre los grupos generales de productos
    for grupo, data_grupo in df_grouped.groupby('Grupo'):
        filas_grupo_temporal = []
        # Agregar una fila especial con el nombre del grupo
        fila_grupo = {col: None for col in df_grouped.columns}
        #print(df_grouped.columns)
        if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
            fila_grupo['CRITERIO PARA AGRUPAR'] = f"**{grupo}**" # Etiqueta de subtotal
        else:
            fila_grupo['DESCRIPCION'] = f"**{grupo}**"

        df_final = pd.concat([df_final, pd.DataFrame([fila_grupo])], ignore_index=True)
        #filas_grupo_temporal.append(fila_grupo)
        print(fila_grupo)
        # Añadir las filas originales para cada grupo
        filas_grupo_temporal.extend(data_grupo.to_dict(orient='records'))

          # Agregar filas con años faltantes con valores en 0
        
        # for _, row in data_grupo.iterrows():
        #     df_final = pd.concat([df_final, row.to_frame().T], ignore_index=True)

        if marcarSubtotal.get() == 1:
            filas_grupo_temporal_acumulada = []  # Lista para almacenar las filas procesadas
            valores_acumulados_por_ano = {}  # Acumuladores por año y nombre común
            nombre_comun_anterior = None
            for _, row in data_grupo.iterrows():
                producto_actual = row[campos_agrupacion_seleccionados[0]]
                
                nombre_comun_actual = obtener_nombre_comun(producto_actual)
                ano = row['Ano']  # Obtener el año actual de la fila

                # Si cambia el nombre común, agregar filas de totales acumulados del anterior
                if nombre_comun_anterior and nombre_comun_actual != nombre_comun_anterior:
                    for ano_acumulado, valores_totales in valores_acumulados_por_ano[nombre_comun_anterior].items():
                        fila_totales = {mes: valores_totales[mes] for mes in valores_totales}
                        fila_totales.update({
                            campos_agrupacion_seleccionados[0]: f"SUBTOTAL: {nombre_comun_anterior}",
                            'Ano': ano_acumulado
                        })
                        if fila_totales["Prom"] > 0:
                            fila_totales["Dur"] = round(fila_totales["Stock"] / fila_totales["Prom"],2)
                        else:
                            fila_totales["Dur"] = 0
                        filas_grupo_temporal_acumulada.append(fila_totales)
                    valores_acumulados_por_ano[nombre_comun_anterior] = {}  # Limpiar acumulador para el siguiente nombre común

                # Inicializar acumulador por año para el nuevo grupo
                if nombre_comun_actual not in valores_acumulados_por_ano:
                    valores_acumulados_por_ano[nombre_comun_actual] = {}
                if ano not in valores_acumulados_por_ano[nombre_comun_actual]:
                    valores_acumulados_por_ano[nombre_comun_actual][ano] = {mes: 0 for mes in ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Total", "12 Meses", "Stock", "Prom", "Dur"]}

                # Sumar valores de la fila actual al acumulador correspondiente
                for mes in valores_acumulados_por_ano[nombre_comun_actual][ano]:
                    cantidad_mes = pd.to_numeric(row[mes], errors='coerce')
                    
                    if pd.notna(cantidad_mes) and cantidad_mes > 0:
                        valores_acumulados_por_ano[nombre_comun_actual][ano][mes] += round(cantidad_mes)

                # Agregar la fila actual al resultado procesado
                filas_grupo_temporal_acumulada.append(row.to_dict())
                nombre_comun_anterior = nombre_comun_actual

            # Agregar filas de totales acumulados del último grupo
            if nombre_comun_anterior:
                for ano_acumulado, valores_totales in valores_acumulados_por_ano[nombre_comun_anterior].items():
                    fila_totales = {mes: valores_totales[mes] for mes in valores_totales}
                    fila_totales.update({
                        campos_agrupacion_seleccionados[0]: f"SUBTOTAL: {nombre_comun_anterior}",
                        'Ano': ano_acumulado
                    })
                    filas_grupo_temporal_acumulada.append(fila_totales)

            # Asegurarse de que las filas procesadas se guarden correctamente
            filas_grupo_temporal = filas_grupo_temporal_acumulada  # Reemplazar las filas originales con las filas procesadas

        for fila in filas_grupo_temporal:
             df_final = pd.concat([df_final, pd.DataFrame([fila])], ignore_index=True)

        # Agregar subtotales por año para el grupo
        for ano in data_grupo['Ano'].unique():
            subtotales_ano = data_grupo[data_grupo['Ano'] == ano].sum(numeric_only=True)
            if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
                subtotales_ano['CRITERIO PARA AGRUPAR'] = f'TOTALES SUB GENERAL'  # Etiqueta de subtotal
            else:
                subtotales_ano['DESCRIPCION'] = f'TOTALES SUB GENERAL'
            subtotales_ano['Ano'] = ano  # Mantener el año en los subtotales

             # Calcular "Dur" como Stock / Prom
            if subtotales_ano["Prom"] > 0:
                subtotales_ano["Dur"] = round(subtotales_ano["Stock"] / subtotales_ano["Prom"],2)
            else:
                subtotales_ano["Dur"] = 0

            df_final = pd.concat([df_final, subtotales_ano.to_frame().T], ignore_index=True)


        # Agregar una fila vacía después de cada grupo (después de los subtotales)
        fila_vacia = {columna: None for columna in df_final.columns}
        df_final = pd.concat([df_final, pd.DataFrame([fila_vacia])], ignore_index=True)

    # Eliminar las columnas innecesarias antes de devolver el DataFrame
    campo_a_ignorar = campos_agrupacion_selec.get()
    df_final = df_final.drop(columns=[campo_a_ignorar], errors='ignore')
    df_final = df_final.drop(columns=['Grupo'], errors='ignore')
    df_final = df_final.drop(columns=['MT2'], errors='ignore')

    # for grupo, data_grupo in df_final.groupby(campos_agrupacion_seleccionados[0]):
    #     for ano in anos_presentes:
    #         if ano not in data_grupo['Ano'].values:
    #             fila_faltante = {col: 0 for col in df_final.columns}
    #             fila_faltante['Ano'] = ano
    #             fila_faltante[campos_agrupacion_seleccionados[0]] = grupo
    #             df_final = pd.concat([df_final, pd.DataFrame([fila_faltante])], ignore_index=True)


    return df_final



def agregar_subtotales_Extra(df, campos_agrupacion_seleccionados):
    """Genera subtotales por grupo de productos similares y agrega subtotales por año."""
    
    # Crear una lista para almacenar todos los resultados
    resultados = []



    # Agrupar por los campos de agrupación seleccionados y Año para obtener las sumas por cada combinación
    df_grouped = df.groupby(campos_agrupacion_seleccionados + ['Ano'], as_index=False).sum(numeric_only=True)

    # Función para obtener el nombre común del producto
    def obtener_nombre_comun(producto):
        match = re.match(r"([A-Za-zÁÉÍÓÚáéíóú0-9]+(?: [A-Za-zÁÉÍÓÚáéíóú0-9]+)*)(?=\s*\d{3}X\d{3})", producto)
        return match.group(1) if match else producto

    # Crear una nueva columna para identificar el grupo general del producto
    campo_agrupacion = campos_agrupacion_selec.get()  # Asumiendo que el usuario seleccionó solo un campo
    if campo_agrupacion == 'MT2':  
        df_grouped['Grupo'] = df_grouped['MT2']
    elif campo_agrupacion == "SUBFA":
        df_grouped['Grupo'] = df_grouped['SUBFA']
    elif campo_agrupacion == "LINEA":
        df_grouped['Grupo'] = df_grouped['LINEA']
    elif campo_agrupacion == "EMPRESA":
        df_grouped['Grupo'] = df_grouped['EMPRESA']
    elif campo_agrupacion == "COMP":
        df_grouped['Grupo'] = df_grouped['COMP']
    elif campo_agrupacion == "numero_bodega":
        df_grouped['Grupo'] = df_grouped['numero_bodega']
    elif campo_agrupacion == "FLIA":
        df_grouped['Grupo'] = df_grouped['FLIA']
    elif campo_agrupacion == "CATEGORIA":
        df_grouped['Grupo'] = df_grouped['CATEGORIA']
    elif campo_agrupacion == "PROVEEDOR":
        df_grouped['Grupo'] = df_grouped['PROVEEDOR']
    elif campo_agrupacion == "LOCAL":
        df_grouped['Grupo'] = df_grouped['LOCAL']
    elif campo_agrupacion == "MEDID":
        df_grouped['Grupo'] = df_grouped['MEDID']
    elif campo_agrupacion == "EMPRESA":
        df_grouped['Grupo'] = df_grouped['EMPRESA']
    else:
        if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
            df_grouped['Grupo'] = df_grouped['CRITERIO PARA AGRUPAR'].apply(obtener_nombre_comun)
        else:
            df_grouped['Grupo'] = df_grouped['DESCRIPCION']

    # Iterar sobre los grupos generales de productos
    for grupo, data_grupo in df_grouped.groupby('Grupo'):
        # Lista temporal para almacenar las filas del grupo actual
        filas_grupo_temporal = []
        
        #AGREGAR FILA DEL NOMBRE DEL GRUPO 
        fila_grupo = {columna: None for columna in df_grouped.columns}
        if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
            #print(campos_agrupacion_seleccionados[0])
            fila_grupo['CRITERIO PARA AGRUPAR'] = f"**{grupo}**"
        else:
            fila_grupo['DESCRIPCION'] = f"**{grupo}**"
        filas_grupo_temporal.append(fila_grupo)
        print(fila_grupo)
        # Añadir las filas originales para cada grupo
        filas_grupo_temporal.extend(data_grupo.to_dict(orient='records'))



        if marcar.get() == 1:
            filas_grupo_temporal_acumulada = []  # Lista para almacenar las filas procesadas
            valores_acumulados_por_ano = {}  # Acumuladores por año y nombre común
            nombre_comun_anterior = None
            filas_grupo_temporal_acumulada.append(fila_grupo)
            for _, row in data_grupo.iterrows():
                producto_actual = row[campos_agrupacion_seleccionados[0]]
                nombre_comun_actual = obtener_nombre_comun(producto_actual)
                ano = row['Ano']  # Obtener el año actual de la fila

                # Si cambia el nombre común, agregar filas de totales acumulados del anterior
                if nombre_comun_anterior and nombre_comun_actual != nombre_comun_anterior:
                    for ano_acumulado, valores_totales in valores_acumulados_por_ano[nombre_comun_anterior].items():
                        fila_totales = {mes: valores_totales[mes] for mes in valores_totales}
                        fila_totales.update({
                            campos_agrupacion_seleccionados[0]: f"TOTALES EXTRA: {nombre_comun_anterior}",
                            'Ano': ano_acumulado
                        })
                        if fila_totales["Prom"] > 0:
                            fila_totales["Dur"] = round(fila_totales["Stock"] / fila_totales["Prom"], 2)
                        else:
                            fila_totales["Dur"] = 0
                        filas_grupo_temporal_acumulada.append(fila_totales)
                    valores_acumulados_por_ano[nombre_comun_anterior] = {}  # Limpiar acumulador para el siguiente nombre común

                # Inicializar acumulador por año para el nuevo grupo
                if nombre_comun_actual not in valores_acumulados_por_ano:
                    valores_acumulados_por_ano[nombre_comun_actual] = {}
                if ano not in valores_acumulados_por_ano[nombre_comun_actual]:
                    valores_acumulados_por_ano[nombre_comun_actual][ano] = {mes: 0 for mes in ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Total", "12 Meses", "Stock", "Prom", "Dur"]}

                # Sumar valores de la fila actual al acumulador correspondiente
                for mes in valores_acumulados_por_ano[nombre_comun_actual][ano]:
                    cantidad_mes = pd.to_numeric(row[mes], errors='coerce')
                    m2 = pd.to_numeric(row['MT2'], errors='coerce')
                    if pd.notna(cantidad_mes) and cantidad_mes > 0 and pd.notna(m2) and m2 > 0:
                        valores_acumulados_por_ano[nombre_comun_actual][ano][mes] += round(m2 * cantidad_mes, 2)

                # Agregar la fila actual al resultado procesado
                filas_grupo_temporal_acumulada.append(row.to_dict())
                nombre_comun_anterior = nombre_comun_actual

            # Agregar filas de totales acumulados del último grupo
            if nombre_comun_anterior:
                for ano_acumulado, valores_totales in valores_acumulados_por_ano[nombre_comun_anterior].items():
                    fila_totales = {mes: valores_totales[mes] for mes in valores_totales}
                    fila_totales.update({
                        campos_agrupacion_seleccionados[0]: f"TOTALES EXTRA: {nombre_comun_anterior}",
                        'Ano': ano_acumulado
                    })
                    if fila_totales["Prom"] > 0:
                        fila_totales["Dur"] = round(fila_totales["Stock"] / fila_totales["Prom"],2)
                    else:
                        fila_totales["Dur"] = 0
                    filas_grupo_temporal_acumulada.append(fila_totales)

            # Asegurarse de que las filas procesadas se guarden correctamente
            filas_grupo_temporal = filas_grupo_temporal_acumulada  # Reemplazar las filas originales con las filas procesadas

        # Agregar subtotales generales por año para el grupo
        for ano in data_grupo['Ano'].unique():
            valores_totales_ano = {mes: 0 for mes in ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Total", "12 Meses", "Stock", "Prom", "Dur"]}

            for _, row in data_grupo[data_grupo['Ano'] == ano].iterrows():
                m2 = pd.to_numeric(row['MT2'], errors='coerce')  
                if pd.notna(m2) and m2 > 0:
                    for mes in valores_totales_ano:
                        cantidad_mes = pd.to_numeric(row[mes], errors='coerce')
                        if pd.notna(cantidad_mes) and cantidad_mes > 0:
                            valor_mes = round(m2 * cantidad_mes, 2)
                            valores_totales_ano[mes] += valor_mes

            # Calcular "Dur" como Stock / Prom
            if valores_totales_ano["Prom"] > 0:
                valores_totales_ano["Dur"] = round(valores_totales_ano["Stock"] / valores_totales_ano["Prom"], 2)
            else:
                valores_totales_ano["Dur"] = 0

            fila_extra = {
                campos_agrupacion_seleccionados[0]: f"TOTALES EXTRA GENERAL",
                'Ano': ano,
            }
            print(fila_extra)
            fila_extra.update(valores_totales_ano)

            # Agregar la fila de totales extra al grupo
            filas_grupo_temporal.append(fila_extra)

        # Agregar una fila vacía completa después del grupo
        filas_grupo_temporal.append({columna: None for columna in filas_grupo_temporal[0].keys()})

        # Agregar las filas del grupo al resultado final
        resultados.extend(filas_grupo_temporal)

    # Devolver el DataFrame final sin la columna 'Grupo'
    df_final = pd.DataFrame(resultados)
    campo_a_ignorar = campos_agrupacion_selec.get()

    df_final = df_final.drop(columns=[campo_a_ignorar], errors='ignore')
    df_final = df_final.drop(columns=['Grupo'], errors='ignore')
    df_final = df_final.drop(columns=['MT2'], errors='ignore')

    return df_final



def Combinar_funciones(df, campos_agrupacion_seleccionados): #fUNCION CON EL CRITERIO DE AGRUPACION
    """Genera subtotales por grupo de productos similares y agrega subtotales por año sin el total general, solo la fila vacía."""
    #print(df)
    
    df_final = pd.DataFrame()
    # Crear una lista para almacenar todos los resultados
    resultados = []

    # Agrupar por los campos de agrupación seleccionados y Año para obtener las sumas por cada combinación
    df_grouped = df.groupby(campos_agrupacion_seleccionados + ['Ano'], as_index=False).sum(numeric_only=True)

    def obtener_nombre_comun(producto):
        match = re.match(r"([A-Za-zÁÉÍÓÚáéíóú0-9]+(?: [A-Za-zÁÉÍÓÚáéíóú0-9]+)*)(?=\s*\d{3}X\d{3})", producto)
        return match.group(1) if match else producto

    # Verificar si el campo seleccionado es el adecuado para agrupar
    campo_agrupacion = campos_agrupacion_selec.get()   # Asumiendo que el usuario seleccionó solo un campo
    if campo_agrupacion == 'MT2':  
        df_grouped['Grupo'] = df_grouped['MT2']
    elif campo_agrupacion == "SUBFA":
        df_grouped['Grupo'] = df_grouped['SUBFA']
    elif campo_agrupacion == "LINEA":
        df_grouped['Grupo'] = df_grouped['LINEA']
    elif campo_agrupacion == "EMPRESA":
        df_grouped['Grupo'] = df_grouped['EMPRESA']
    elif campo_agrupacion == "COMP":
        df_grouped['Grupo'] = df_grouped['COMP']
    elif campo_agrupacion == "numero_bodega":
        df_grouped['Grupo'] = df_grouped['numero_bodega']
    elif campo_agrupacion == "FLIA":
        df_grouped['Grupo'] = df_grouped['FLIA']
    elif campo_agrupacion == "CATEGORIA":
        df_grouped['Grupo'] = df_grouped['CATEGORIA']
    elif campo_agrupacion == "PROVEEDOR":
        df_grouped['Grupo'] = df_grouped['PROVEEDOR']
    elif campo_agrupacion == "LOCAL":
        df_grouped['Grupo'] = df_grouped['LOCAL']
    elif campo_agrupacion == "MEDID":
        df_grouped['Grupo'] = df_grouped['MEDID']
    elif campo_agrupacion == "EMPRESA":
        df_grouped['Grupo'] = df_grouped['EMPRESA']
    else:
        if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
            df_grouped['Grupo'] = df_grouped['CRITERIO PARA AGRUPAR'].apply(obtener_nombre_comun)
        else:
            df_grouped['Grupo'] = df_grouped['DESCRIPCION']

    # Iterar sobre los grupos generales de productos
    for grupo, data_grupo in df_grouped.groupby('Grupo'):
        # Lista temporal para almacenar las filas del grupo actual
        filas_grupo_temporal = []

          # === AGREGAR FILA DEL NOMBRE DEL GRUPO ===
        fila_grupo = {columna: None for columna in df_grouped.columns}
        if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
            fila_grupo['CRITERIO PARA AGRUPAR'] = f"**{grupo}**"
        else:
            fila_grupo['DESCRIPCION'] = f"**{grupo}**"
        filas_grupo_temporal.append(fila_grupo)
        # Añadir las filas originales para cada grupo
        filas_grupo_temporal.extend(data_grupo.to_dict(orient='records'))
        
        if marcar.get() == 1:
            filas_grupo_temporal_acumulada = []  # Lista para almacenar las filas procesadas
            valores_acumulados_por_ano = {}  # Acumuladores por año y nombre común
            nombre_comun_anterior = None
            filas_grupo_temporal_acumulada.append(fila_grupo)
            for _, row in data_grupo.iterrows():
                producto_actual = row[campos_agrupacion_seleccionados[0]]
                nombre_comun_actual = obtener_nombre_comun(producto_actual)
                ano = row['Ano']  # Obtener el año actual de la fila

                # Si cambia el nombre común, agregar filas de totales acumulados del anterior
                if nombre_comun_anterior and nombre_comun_actual != nombre_comun_anterior:
                    for ano_acumulado, valores_totales in valores_acumulados_por_ano[nombre_comun_anterior].items():
                        fila_totales = {mes: valores_totales[mes] for mes in valores_totales}
                        fila_totales.update({
                            campos_agrupacion_seleccionados[0]: f"TOTALES EXTRA: {nombre_comun_anterior}",
                            'Ano': ano_acumulado
                        })
                        if fila_totales["Prom"] > 0:
                            fila_totales["Dur"] = round(fila_totales["Stock"] / fila_totales["Prom"], 2)
                        else:
                            fila_totales["Dur"] = 0
                        filas_grupo_temporal_acumulada.append(fila_totales)
                    valores_acumulados_por_ano[nombre_comun_anterior] = {}  # Limpiar acumulador para el siguiente nombre común

                # Inicializar acumulador por año para el nuevo grupo
                if nombre_comun_actual not in valores_acumulados_por_ano:
                    valores_acumulados_por_ano[nombre_comun_actual] = {}
                if ano not in valores_acumulados_por_ano[nombre_comun_actual]:
                    valores_acumulados_por_ano[nombre_comun_actual][ano] = {mes: 0 for mes in ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Total", "12 Meses", "Stock", "Prom", "Dur"]}

                # Sumar valores de la fila actual al acumulador correspondiente
                for mes in valores_acumulados_por_ano[nombre_comun_actual][ano]:
                    cantidad_mes = pd.to_numeric(row[mes], errors='coerce')
                    m2 = pd.to_numeric(row['MT2'], errors='coerce')
                    if pd.notna(cantidad_mes) and cantidad_mes > 0 and pd.notna(m2) and m2 > 0:
                        valores_acumulados_por_ano[nombre_comun_actual][ano][mes] += round(m2 * cantidad_mes,2)

                # Agregar la fila actual al resultado procesado
                filas_grupo_temporal_acumulada.append(row.to_dict())
                nombre_comun_anterior = nombre_comun_actual

            # Agregar filas de totales acumulados del último grupo
            if nombre_comun_anterior:
                for ano_acumulado, valores_totales in valores_acumulados_por_ano[nombre_comun_anterior].items():
                    fila_totales = {mes: valores_totales[mes] for mes in valores_totales}
                    fila_totales.update({
                        campos_agrupacion_seleccionados[0]: f"TOTALES EXTRA: {nombre_comun_anterior}",
                        'Ano': ano_acumulado
                    })
                    if fila_totales["Prom"] > 0:
                        fila_totales["Dur"] = round(fila_totales["Stock"] / fila_totales["Prom"], 2)
                    else:
                        fila_totales["Dur"] = 0
                    filas_grupo_temporal_acumulada.append(fila_totales)

            # Asegurarse de que las filas procesadas se guarden correctamente
            filas_grupo_temporal = filas_grupo_temporal_acumulada  # Reemplazar las filas originales con las filas procesadas

        # Agregar subtotales por año para el grupo
        for ano in data_grupo['Ano'].unique():
            subtotales_ano = data_grupo[data_grupo['Ano'] == ano].sum(numeric_only=True)
            if campos_agrupacion_seleccionados[0] == 'CRITERIO PARA AGRUPAR':
                subtotales_ano['CRITERIO PARA AGRUPAR'] = f'TOTALES SUB'  # Etiqueta de subtotal
            else:
                subtotales_ano['DESCRIPCION'] = f'TOTALES SUB GENERAL'

            subtotales_ano['Ano'] = ano
            # Calcular "Dur" como Stock / Prom
            if subtotales_ano["Prom"] > 0:
                subtotales_ano["Dur"] = round(subtotales_ano["Stock"] / subtotales_ano["Prom"], 2)
            else:
                subtotales_ano["Dur"] = 0
            filas_grupo_temporal.append(subtotales_ano.to_dict())
            
    
        # Agregar las filas "Extra" (agrupando por Grupo)
        for ano in data_grupo['Ano'].unique():
            # Inicializar los valores de los meses a 0
            valores_totales_ano = {mes: 0 for mes in ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sept", "Oct", "Nov", "Dic", "Total", "12 Meses", "Stock", "Prom", "Dur"]}

            # Iterar sobre las filas para este año específico
            for _, row in data_grupo[data_grupo['Ano'] == ano].iterrows():
                if 'MT2' in row:  # Verificar si la columna 'MT2' existe
                    m2 = pd.to_numeric(row['MT2'], errors='coerce')  # Obtener el valor ya calculado del m²
                    if pd.notna(m2) and m2 > 0:
                        for mes in valores_totales_ano:
                            cantidad_mes = pd.to_numeric(row[mes], errors='coerce')
                            if pd.notna(cantidad_mes) and cantidad_mes > 0:
                                valor_mes = round(m2 * cantidad_mes,2)
                                valores_totales_ano[mes] += valor_mes

            # Calcular "Dur" como Stock / Prom
            if valores_totales_ano["Prom"] > 0:
                valores_totales_ano["Dur"] = round(valores_totales_ano["Stock"] / valores_totales_ano["Prom"], 2)
            else:
                valores_totales_ano["Dur"] = 0

            fila_extra = {
                campos_agrupacion_seleccionados[0] : f"TOTALES EXTRA GENERAL",
                'Ano': ano,
            }
            fila_extra.update(valores_totales_ano)
            
            # Agregar la fila de totales extra al grupo
            filas_grupo_temporal.append(fila_extra)

        # Agregar una fila vacía completa después del grupo
        filas_grupo_temporal.append({columna: None for columna in filas_grupo_temporal[0].keys()})

        # Agregar las filas del grupo al resultado final
        resultados.extend(filas_grupo_temporal)

    # Devolver el DataFrame final sin la columna 'Grupo'
    df_final = pd.DataFrame(resultados)
    campo_a_ignorar = campos_agrupacion_selec.get()
    
    df_final = df_final.drop(columns=[campo_a_ignorar], errors='ignore')
    df_final = df_final.drop(columns=['Grupo'], errors='ignore')
    df_final = df_final.drop(columns=['MT2'], errors='ignore')

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
                       "Jul", "Ago", "Sept", "Oct", "Nov", "Dic","Código","EMPRESA",
                       "numero_bodega","LOCAL","MT2","PROVEEDOR","COMP","LINEA","CATEGORIA","FLIA","SUBFA", "MEDID"]
campos_s_filtro = ["Ano", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                       "Jul", "Ago", "Sept", "Oct", "Nov", "Dic"]

campos_filtrados = [campo for campo in campos if campo not in campos_obligatorios]
campos_sin_filtrar = [campo for campo in campos if campo not in campos_s_filtro]
campos_seleccionados = {}
campos_agrupacion = {}
campos_agrupacion_selec = {}
# Estilo de ttk
style = ttk.Style()
style.theme_use('clam')
style.configure('TButton',
                font=('Segoe UI', 10),
                padding=5,
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


# Frame para la selección de campos 
frame_seleccion = ttk.LabelFrame(ventana, text="Selecciona los campos", padding=(1, 2), style="TFrame")
frame_seleccion.grid(row=0, column=0, rowspan=2, sticky="n", padx=(1, 1), pady=10)  # Ajuste para mover a la izquierda

# Crear y mostrar los checkbuttons de selección de campos
for i, campo in enumerate(campos_filtrados):
    var = tk.IntVar()
    checkbutton = ttk.Checkbutton(frame_seleccion, text=campo, variable=var, style="TCheckbutton")
    checkbutton.pack(side="top", anchor="w", padx=2, pady=2)
    campos_seleccionados[campo] = var
    campos_agrupacion[campo] = var

    

# Frame para los filtros
frame_filtro = ttk.LabelFrame(ventana, text="Filtrar por", padding=(10, 5), style="TFrame")
frame_filtro.grid(row=0, column=2, sticky="n", padx=10, pady=10)

# Dropdown para seleccionar el campo a filtrar
combo_filtro_campo = ttk.Combobox(frame_filtro, values=campos_sin_filtrar)
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


datos_en_memoria = []

# Función para cargar los datos del campo seleccionado a la memoria
def cargar_datos_en_memoria():
    """Carga los datos del campo seleccionado en una lista."""
    global datos_en_memoria
    campo = combo_filtro_campo.get()  # Campo seleccionado para filtrar
    if campo and campo in campos_sin_filtrar:
        df = obtener_datos("Estadistica")  # Obtén los datos del campo seleccionado
        datos_en_memoria = list(df[campo].dropna().unique())  # Guarda los valores únicos en memoria
        combo_filtro_valor['values'] = datos_en_memoria  # Actualiza el dropdown con todos los valores

# Función para filtrar valores en la lista en memoria
def filtrar_valores_en_memoria(event=None):
    """Filtra los valores del dropdown usando los datos en memoria mientras el usuario escribe."""
    texto = combo_filtro_valor.get().upper() # Obtén el texto escrito por el usuario
    
    if datos_en_memoria:  # Solo filtra si hay datos cargados
        # Filtra los valores en memoria según el texto ingresado
        valores_filtrados = [str(valor) for valor in datos_en_memoria if texto in str(valor).upper()]
        combo_filtro_valor['values'] = valores_filtrados  # Actualiza el dropdown con los valores filtrados
        
        # Cierra el menú desplegable si está abierto
        combo_filtro_valor.event_generate('<Escape>')  # Simula presionar Escape para cerrar el menú

        # Mantén el texto escrito y reposiciona el cursor
        combo_filtro_valor.delete(0, 'end')  # Borra el contenido actual del campo
        combo_filtro_valor.insert(0, texto)  # Inserta el texto escrito por el usuario
        combo_filtro_valor.icursor(len(texto))  # Posiciona el cursor al final del texto

# Función para abrir el menú al presionar Enter
def abrir_menu_con_enter(event=None):
    """Despliega el menú del Combobox cuando el usuario presiona Enter."""
    texto = combo_filtro_valor.get().upper()
    if texto:  # Solo abre el menú si hay texto escrito
        combo_filtro_valor.event_generate('<Down>')  # Abre el menú del combobox

# Vincular el evento de escritura en el campo de valores
combo_filtro_valor.bind("<KeyRelease>", filtrar_valores_en_memoria)

# Vincular el evento de presionar Enter para abrir el menú
combo_filtro_valor.bind("<Return>", abrir_menu_con_enter)

# Vincular la carga de datos en memoria al seleccionar un campo
combo_filtro_campo.bind("<<ComboboxSelected>>", lambda e: cargar_datos_en_memoria())



frame_agrupacion = ttk.LabelFrame(ventana, text="Agrupar por", padding=(10, 5), style="TFrame")
frame_agrupacion.grid(row=0, column=1, sticky="n", padx=10, pady=10)

# Usamos una variable Tkinter para almacenar el campo seleccionado 
campo_agrupacion_actual_seleccionado = tk.StringVar()

# Crear los RadioButtons para la selección de agrupación
for i, campo in enumerate(campos_sin_filtrar):
    radio_button = ttk.Radiobutton(frame_agrupacion, text=campo, variable=campo_agrupacion_actual_seleccionado, value=campo, style="TRadiobutton")
    radio_button.pack(side="top", anchor="w", padx=2, pady=2)

# Esta variable almacenará el campo seleccionado para agrupación
campos_agrupacion_selec = campo_agrupacion_actual_seleccionado
#print(campo_agrupacion_actual_seleccionado.get())

# Variable para incluir subtotales agrupados
agregar_subtotales_var = tk.IntVar()
check_subtotales = tk.Checkbutton(ventana, text="Incluir SubT General", variable=agregar_subtotales_var)
check_subtotales.grid(row=1, column=0, padx=(60,10), pady=5, sticky="w")  # Alineado a la izquierda (west)

# Variable para incluir subtotales extra
agregar_subtotales_Ext = tk.IntVar()
check_subtotales_extra = tk.Checkbutton(ventana, text="Incluir subT General extra MT2", variable=agregar_subtotales_Ext)
check_subtotales_extra.grid(row=4, column=0, padx=(60,10), pady=5, sticky="w")  # Alineado a la izquierda


marcar = tk.IntVar()
check_marcar = tk.Checkbutton(ventana, text="Incluir agrupacion extra MT2", variable=marcar)
check_marcar.grid(row=5, column=0, padx=(60,10), pady=5, sticky="w")  # Alineado a la izquierda

marcarSubtotal = tk.IntVar()
check_marcar = tk.Checkbutton(ventana, text="Incluir agrupacion subtotal", variable=marcarSubtotal)
check_marcar.grid(row=6, column=0, padx=(60,10), pady=5, sticky="w")  # Alineado a la izquierda


ventana.grid_columnconfigure(0, weight=1)  
ventana.grid_columnconfigure(1, weight=1)  
ventana.grid_columnconfigure(2, weight=1)  

boton_generar = tk.Button(ventana, text="Generar informe", command=generar_reporte)
boton_generar.grid(row=7, column=1, pady=5, sticky="ew")


# Ejecutar la aplicación
ventana.mainloop()



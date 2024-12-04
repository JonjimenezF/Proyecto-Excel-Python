
# Generador de Reportes de Estadísticas de Compras

Este programa es una herramienta diseñada para la generación de reportes basados en datos almacenados en una base de datos Access. Permite aplicar filtros avanzados, realizar cálculos estadísticos, manejar información de stock desde un archivo Excel y generar reportes personalizados en formato Excel.

## **Características**
- Conexión a una base de datos Access para consultar y manipular datos.
- Aplicación de filtros dinámicos en tiempo real.
- Cálculo de estadísticas mensuales y anuales.
- Integración con un archivo Excel para gestionar datos de stock.
- Generación de reportes agrupados y personalizados en formato Excel.

## **Requisitos del sistema**

### **Dependencias**
Antes de ejecutar el programa, asegúrate de cumplir con los siguientes requisitos:

1. **Software:**
   - Python 3.8 o superior
   - Controlador de Microsoft Access adecuado para tu sistema operativo (32 o 64 bits)

2. **Librerías de Python:**  
   Instala las siguientes librerías con el comando:
   ```bash
   pip install pandas pyodbc openpyxl
   ```

### **Archivos necesarios**
Coloca los siguientes archivos en la misma carpeta que el programa principal:
- `BASE.accdb`: Archivo de base de datos Access.
- `Tabla dinámica Analysis (x_bi_sql_view.stock_bodega_sc).xlsm`: Archivo Excel con información de stock.

## **Instrucciones de uso**

### **1. Clonar el repositorio**
Clona o descarga este repositorio en tu máquina local:
```bash
git clone <url_del_repositorio>
```

### **2. Configurar los archivos**
- Coloca el archivo `BASE.accdb` en la carpeta `estadistica_compra`.
- Coloca el archivo Excel `Tabla dinámica Analysis (x_bi_sql_view.stock_bodega_sc).xlsm` en la misma carpeta.

### **3. Ejecutar el programa**
Para ejecutar el programa, utiliza el siguiente comando en tu terminal:
```bash
python <nombre_del_archivo>.py
```

## **Interfaz gráfica**
- Selecciona los campos deseados para incluir en el reporte.
- Aplica filtros según sea necesario.
- Genera el reporte en formato Excel. Este se guardará como `reporte_calculado.xlsx`.

## **Funcionalidades principales**
- **Filtros dinámicos:** Permite seleccionar y eliminar filtros directamente desde la interfaz gráfica.
- **Cálculo de métricas clave:** Genera métricas como promedios mensuales y duración del inventario.
- **Agrupación personalizada:** Agrupa los datos por campos seleccionados por el usuario.
- **Exportación de reportes:** Exporta reportes detallados en formato Excel con cálculos automáticos.

## **Archivos generados**
- `reporte_calculado.xlsx`: Archivo Excel que contiene los resultados tras aplicar los filtros y cálculos.

## **Notas adicionales**
- Asegúrate de instalar el controlador de Access adecuado según tu sistema operativo.
- Verifica que todos los archivos requeridos estén en la ubicación correcta antes de ejecutar el programa.
- Si encuentras algún error, revisa las dependencias y la configuración de los archivos.

"""
Este script extrae tablas de un estado de cuenta en PDF de Bancolombia.

Dependencias requeridas:
- camelot-py[cv]: Para la extracción de tablas de PDFs.
- pandas: Para la manipulación de los datos extraídos.
- matplotlib: Para visualizar las áreas de extracción de tablas.
- ghostscript: Requerido por Camelot para procesar PDFs.

"""

import camelot
import pandas as pd
import os
import matplotlib.pyplot as plt
from pypdf import PdfReader

def formatear_numero(numero):
  """
  Formatea un número con separadores de miles y dos decimales.

  Args:
    numero: El número a formatear (puede ser un int, float o un string numérico).

  Returns:
    Una cadena de texto con el número formateado.
  """
  return f"{numero:,.2f}"

def extract_table(pdf_path, page, table_area, columns, title="Tabla", visualize=False):
    """
    Extrae una tabla específica de una página de un archivo PDF.

    Args:
        pdf_path (str): Ruta al archivo PDF.
        page (str): Número de la página de la cual extraer la tabla (como string).
        table_area (list): Lista con una cadena que define el área de la tabla (ej. ['x1,y1,x2,y2']).
        columns (list): Lista con una cadena de las posiciones de las columnas (ej. ['c1,c2,c3...']).
        title (str): Título para la visualización del gráfico.
        visualize (bool): Si es True, muestra un gráfico de la extracción.

    Returns:
        pd.DataFrame: Un DataFrame de pandas con la tabla extraída.
                      Retorna None si no se encuentra ninguna tabla.
    """
    try:
        tables = camelot.read_pdf(
            pdf_path,
            flavor='stream',
            pages=page,
            table_areas=table_area,
            columns=columns
        )
        if tables.n > 0:
            if visualize:
                print(f"Generando visualización para: {title}")
                # CORRECCIÓN: Se pasa la primera tabla (tables[0]) a la función de ploteo, no la lista de tablas.
                # 'all' muestra el texto, las líneas y los contornos de la tabla.
                plot = camelot.plot(tables[0], kind='grid')
                plt.suptitle(title) # Usar suptitle para que no se solape con el título de camelot
                plt.show()
            return tables[0].df
        else:
            print(f"Advertencia: No se encontró ninguna tabla en la página {page} con el área {table_area}")
            return None
    except Exception as e:
        print(f"Ocurrió un error extrayendo la tabla de la página {page}: {e}")
        return None

def process_pdf(activar_visualizacion = False,pdf_path=""):
    """
    Función principal que orquesta la extracción de tablas del PDF.
    """
    # --- CONFIGURACIÓN ---
    if pdf_path == "":
      raise Exception("Debe especificar la ruta del archivo PDF.")
    reader = PdfReader(pdf_path)

    number_of_pages = len(reader.pages)
    print(f"El archivo '{pdf_path}' tiene {number_of_pages} páginas.")
    if not os.path.exists(pdf_path):
        print(f"Error: El archivo '{pdf_path}' no se encontró.")
        return

    # --- Configuraciones para cada tabla ---
    bancolombia_config_resumen = {
        "title": "Tabla de Resumen",
        "page": '1',
        "area": ['0,450,600,505'],
        "columns": ['110,190,300,410,490']
    }
    bancolombia_config_movimientos = {
        "title": "Tabla de Movimientos",
        "page": '1',
        "area": ['0,70,600,430'],
        "columns": ['90,280,350,410,500']
    }

    bancolombia_config_pg_mayor_a_1 = {
        "title": "Tabla de Movimientos",
        "page": 'all',
        "area": ['0,70,600,640'],
        "columns": ['90,280,350,410,500']
    }
    templates=[]
    print(number_of_pages,"///")
    for pag in range(number_of_pages):
      if pag >= 1:
        bancolombia_config_pg_mayor_a_1 = {
        "title": "Tabla de Movimientos",
        "page": str(pag+1),
        "area": ['0,70,600,610'],
        "columns": ['90,280,350,410,500']
    }
        templates.append(bancolombia_config_pg_mayor_a_1)
        if activar_visualizacion:
          print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$",templates)

    print(f"Iniciando extracción de tablas del archivo: {pdf_path}")
    if activar_visualizacion:
        print("La visualización está ACTIVADA. Se mostrarán gráficos de depuración.")

    # --- Extracción de la tabla de Resumen ---
    print(f"\n--- Extrayendo {bancolombia_config_resumen['title']} ---")
    df_resumen = extract_table(
        pdf_path,
        bancolombia_config_resumen["page"],
        bancolombia_config_resumen["area"],
        bancolombia_config_resumen["columns"],
        title=bancolombia_config_resumen["title"],
        visualize=activar_visualizacion
    )

    if df_resumen is not None and activar_visualizacion:
        print(f"¡{bancolombia_config_resumen['title']} extraída con éxito!")
        print(df_resumen.to_string())

    # --- Extracción de la tabla de Movimientos ---
    if activar_visualizacion:
      print(f"\n--- Extrayendo {bancolombia_config_movimientos['title']} ---")
    df_movimientos = extract_table(
        pdf_path,
        bancolombia_config_movimientos["page"],
        bancolombia_config_movimientos["area"],
        bancolombia_config_movimientos["columns"],
        title=bancolombia_config_movimientos["title"],
        visualize=activar_visualizacion
    )
    if df_movimientos is not None and activar_visualizacion:
        print(f"¡{bancolombia_config_movimientos['title']} extraída con éxito!")
        print(df_movimientos.to_string())

# --- Extracción de la tabla de movimientos mayor a uno ---
    if activar_visualizacion:
      print(f"\n--- Extrayendo {bancolombia_config_pg_mayor_a_1['title']} ---")
    movimientos_list = []
    for template in range(len(templates)):
      df_pg_mayor_a_1 = extract_table(
          pdf_path,
          templates[template]["page"],
          templates[template]["area"],
          templates[template]["columns"],
          title=templates[template]["title"],
          visualize=activar_visualizacion
      )
      movimientos_list.append(df_pg_mayor_a_1)
      #df_movimientos=pd.concat([df_movimientos,df_pg_mayor_a_1])
      if df_pg_mayor_a_1 is not None and activar_visualizacion:
          print(f"¡{bancolombia_config_pg_mayor_a_1['title']} extraída con éxito!")
          print(df_pg_mayor_a_1.to_string())
      df_movimientos = pd.concat(movimientos_list, ignore_index=True)
############################################################################# retrive data
    try:
          datosfinales=df_resumen.set_index(0)[2].to_dict()

          df = df_movimientos.iloc[1:] # Eliminamos la primera fila de encabezado que no tiene valor
          valor_maximo_absoluto = pd.to_numeric(df[4].str.replace(',', ''), errors='coerce').abs().max()
          datosfinales["Valor de movimiento maximo en el mes en pesos"]=formatear_numero(valor_maximo_absoluto)
          dataf_2_excel=pd.DataFrame(datosfinales,index=[0])
          dataf_2_excel.to_excel("movimientos.xlsx", index=False)
          print("######## Resultado de extracción ########", datosfinales)
          #df.to_csv("movimientos.csv")
          return datosfinales
    except Exception as e:
          print("No se pudo extraer la Tabla 1 (Resumen) con las coordenadas dadas.",e)
          return {"error": f"No se pudo extraer la Tabla 1 (Resumen) con las coordenadas dadas. {e}"}
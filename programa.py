import pandas as pd
import tkinter as tk
from tkinter import filedialog, StringVar

nombre_columna_fecha = 'Fecha-inicio'
nombre_columna_tipo = 'Unnamed: 1'
nombre_columna_ws = 'WS'
nombre_columna_entorno = 'Entorno'
nombre_columna_numero_ticket = 'N°TK'
nombre_columna_estado = 'Estado'  

pd.set_option('display.max_colwidth', None)

def mostrar_conteo_por_mes(df, columna_fecha, tipo):
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce', format='%d-%b', dayfirst=True)
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce', format='%d/%m/%Y')

    df['Mes'] = df[columna_fecha].dt.strftime('%Y-%m')

    df_tipo = df[df[nombre_columna_tipo].str.contains(tipo, case=False, na=False)].copy()
    df_tipo['Informacion'] = df_tipo.apply(lambda x: f"Tipo: {x[nombre_columna_tipo]}, Número de Ticket: {x[nombre_columna_numero_ticket]}, WS: {x[nombre_columna_ws]}, Entorno: {x[nombre_columna_entorno]}, Estado: {x[nombre_columna_estado]}", axis=1)

    # Eliminar duplicados antes de realizar el conteo
    df_tipo = df_tipo.drop_duplicates(subset=['Mes', 'Informacion'])

    # Obtener el total de estados y WS por mes
    total_estados_por_mes = df_tipo.groupby(['Mes', 'Estado']).size().reset_index(name='Total Estados')
    total_ws_por_mes = df_tipo.groupby(['Mes', nombre_columna_ws]).size().reset_index(name='Total WS')

    conteo_total_por_mes = df_tipo.groupby(['Mes', 'Informacion']).size().reset_index(name='Count')

    return conteo_total_por_mes, total_estados_por_mes, total_ws_por_mes

def mostrar_conteo_total_por_mes(df, columna_fecha):
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce', format='%d-%b', dayfirst=True)
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce', format='%d/%m/%Y')

    df['Mes'] = df[columna_fecha].dt.strftime('%Y-%m')

    df_requerimientos, total_estados_requerimientos, total_ws_requerimientos = mostrar_conteo_por_mes(df, columna_fecha, 'Requerimiento')
    df_incidentes, total_estados_incidentes, total_ws_incidentes = mostrar_conteo_por_mes(df, columna_fecha, 'Incidente')

    total_requerimientos_por_mes = df_requerimientos.groupby('Mes')['Count'].sum().reset_index(name='Total Requerimientos')
    total_incidentes_por_mes = df_incidentes.groupby('Mes')['Count'].sum().reset_index(name='Total Incidentes')

    return df_requerimientos, df_incidentes, total_requerimientos_por_mes, total_incidentes_por_mes, total_estados_requerimientos, total_estados_incidentes, total_ws_requerimientos, total_ws_incidentes

def cargar_excel():
    global df_requerimientos, df_incidentes, total_requerimientos, total_incidentes, total_estados_requerimientos, total_estados_incidentes, total_ws_requerimientos, total_ws_incidentes

    ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])

    if ruta_archivo:
        # Leer todas las hojas del archivo
        xls = pd.ExcelFile(ruta_archivo)
        
        # Inicializar los DataFrames
        df_requerimientos = pd.DataFrame()
        df_incidentes = pd.DataFrame()
        total_requerimientos = pd.DataFrame()
        total_incidentes = pd.DataFrame()
        total_estados_requerimientos = pd.DataFrame()
        total_estados_incidentes = pd.DataFrame()
        total_ws_requerimientos = pd.DataFrame()
        total_ws_incidentes = pd.DataFrame()

        # Procesar cada hoja del archivo
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, header=0)

            if nombre_columna_fecha in df.columns and nombre_columna_tipo in df.columns and nombre_columna_estado in df.columns:
                # Obtener los resultados por mes para requerimientos e incidentes
                df_req, df_inc, total_req, total_inc, total_estados_req, total_estados_inc, total_ws_req, total_ws_inc = mostrar_conteo_total_por_mes(df, nombre_columna_fecha)

                # Concatenar los resultados
                df_requerimientos = pd.concat([df_requerimientos, df_req], ignore_index=True)
                df_incidentes = pd.concat([df_incidentes, df_inc], ignore_index=True)
                total_requerimientos = pd.concat([total_requerimientos, total_req], ignore_index=True)
                total_incidentes = pd.concat([total_incidentes, total_inc], ignore_index=True)
                total_estados_requerimientos = pd.concat([total_estados_requerimientos, total_estados_req], ignore_index=True)
                total_estados_incidentes = pd.concat([total_estados_incidentes, total_estados_inc], ignore_index=True)
                total_ws_requerimientos = pd.concat([total_ws_requerimientos, total_ws_req], ignore_index=True)
                total_ws_incidentes = pd.concat([total_ws_incidentes, total_ws_inc], ignore_index=True)

        # Mostrar resultados de lo que se selecciona
        mes_seleccionado = var_mes.get()
        tipo_seleccionado = var_tipo.get()

        if tipo_seleccionado == 'Requerimientos':
            df_resultado = df_requerimientos[df_requerimientos['Mes'] == mes_seleccionado]
            total_por_mes = total_requerimientos[total_requerimientos['Mes'] == mes_seleccionado]['Total Requerimientos'].values
            total_por_mes = total_por_mes[0] if len(total_por_mes) > 0 else 0
            total_por_mes_estados = total_estados_requerimientos[total_estados_requerimientos['Mes'] == mes_seleccionado].groupby('Estado')['Total Estados'].sum().reset_index()
            total_por_mes_ws = total_ws_requerimientos[total_ws_requerimientos['Mes'] == mes_seleccionado].groupby(nombre_columna_ws)['Total WS'].sum().reset_index()
        else:
            df_resultado = df_incidentes[df_incidentes['Mes'] == mes_seleccionado]
            total_por_mes = total_incidentes[total_incidentes['Mes'] == mes_seleccionado]['Total Incidentes'].values
            total_por_mes = total_por_mes[0] if len(total_por_mes) > 0 else 0
            total_por_mes_estados = total_estados_incidentes[total_estados_incidentes['Mes'] == mes_seleccionado].groupby('Estado')['Total Estados'].sum().reset_index()
            total_por_mes_ws = total_ws_incidentes[total_ws_incidentes['Mes'] == mes_seleccionado].groupby(nombre_columna_ws)['Total WS'].sum().reset_index()

        print(f"\nResultados del mes {mes_seleccionado} para {tipo_seleccionado}:")
        if not df_resultado.empty:
            print(df_resultado[['Mes', 'Informacion']])
        else:
            print(f"No se encontraron {tipo_seleccionado} para el mes {mes_seleccionado}.")

        print(f"\nTotal de {tipo_seleccionado} por Mes:")
        print(f"{mes_seleccionado}: {total_por_mes}")

        print(f"\nResueltos o en proceso por Mes:")
        print(f"{mes_seleccionado}:")
        print(total_por_mes_estados)

        print(f"\nTotal de WS por Mes:")
        print(f"{mes_seleccionado}:")
        print(total_por_mes_ws)

    else:
        print("Las columnas requeridas no se encontraron en el DataFrame.")

def mostrar_totales_acumulados():
    global total_requerimientos, total_incidentes, total_estados_requerimientos, total_estados_incidentes, total_ws_requerimientos, total_ws_incidentes

    # Verificar si los totales están definidos
    if 'total_requerimientos' not in globals() or 'total_incidentes' not in globals() or 'total_estados_requerimientos' not in globals() or 'total_estados_incidentes' not in globals() or 'total_ws_requerimientos' not in globals() or 'total_ws_incidentes' not in globals():
        print("No se han cargado datos para mostrar totales acumulados.")
        return

    # Mostrar totales acumulados para requerimientos e incidentes
    total_acumulado_requerimientos = total_requerimientos['Total Requerimientos'].sum()
    total_acumulado_incidentes = total_incidentes['Total Incidentes'].sum()

    # Mostrar totales por mes
    print("\nTotales por Mes - Requerimientos:")
    print(total_requerimientos)

    print("\nTotales por Mes - Incidentes:")
    print(total_incidentes)

    print("\nResueltos o en proceso por Mes - Requerimientos:")
    print(total_estados_requerimientos)

    print("\nResueltos o en proceso por Mes - Incidentes:")
    print(total_estados_incidentes)

    print("\nTotal de WS por Mes - Requerimientos:")
    print(total_ws_requerimientos)

    print("\nTotal de WS por Mes - Incidentes:")
    print(total_ws_incidentes)

# Crear un diccionario para almacenar los resultados por mes para requerimientos e incidentes
resultados_por_mes_requerimientos = {}
resultados_por_mes_incidentes = {}

app = tk.Tk()
app.title("Programa")

# Variables para almacenar la selección del usuario
var_mes = StringVar()
var_tipo = StringVar()

# Botones de opción para seleccionar el mes
lbl_mes = tk.Label(app, text="Selecciona el Mes:")
lbl_mes.pack()
meses = ["2023-05", "2023-06", "2023-07", "2023-08", "2023-09", "2023-10", "2023-11", "2023-12", "2024-01"]  
opcion_mes = tk.OptionMenu(app, var_mes, *meses)
opcion_mes.pack()

# Botones de opción para seleccionar el tipo de información
lbl_tipo = tk.Label(app, text="Selecciona el Tipo de Información:")
lbl_tipo.pack()
tipos = ["Requerimientos", "Incidentes"]
opcion_tipo = tk.OptionMenu(app, var_tipo, *tipos)
opcion_tipo.pack()

# Botón para cargar el archivo Excel
btn_cargar_excel = tk.Button(app, text="Cargar Excel", command=cargar_excel)
btn_cargar_excel.pack(pady=10)

# Botón para mostrar totales acumulados
btn_totales_acumulados = tk.Button(app, text="Mostrar Total Por Mes", command=mostrar_totales_acumulados)
btn_totales_acumulados.pack(pady=10)

app.mainloop()

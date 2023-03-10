# Librerias
import os
import pandas as pd
import time
import openpyxl

# Variables globales ( LAS DIRECCIONES DE LAS CARPETAS SERAN VERIFICADAS POR EL NUMERO DE SEMANA ADMINISTRATIVA, asi solo nos preocupamos por cambiar el valor de "week_number")

folder_path = '/home/xeroxv23/Documents/Proyectos GCPI/acumulados_sueldos_semanales/SEMANA_{}'
week_number = 4
final_path = folder_path.format(week_number)
acugen_path = '/home/xeroxv23/Documents/Proyectos GCPI/acumulados_sueldos_semanales/ACUGEN_SEM_{}/acumulado_sueldos_sem{}.csv'
acugen_final = acugen_path.format(week_number, week_number)
acugen_excel = "/home/xeroxv23/Documents/Proyectos GCPI/acumulados_sueldos_semanales/ACUGEN_SEM_{}/nuevo_acugen_sem{}.xlsx"
acugen_excel_final = acugen_excel.format(week_number, week_number)

# contador de tiempo
start_time = time.time()

# especifica la ruta de la carpeta
path = final_path

# obtiene una lista de todos los archivos xlsm en la carpeta
files = [f for f in os.listdir(path) if f.endswith('.xlsm')]

# inicializa una lista vacía para almacenar los dataframes individuales
df_list = []

# itera a través de cada archivo xlsm y lee los datos en un dataframe solo si la celda del archivo de excel coincide la celda B10 con el numero de semana
for file in files:
    df_celda_semana = pd.read_excel(os.path.join(path, file),sheet_name=0, engine='openpyxl', header=None, skiprows=9, nrows=1, usecols=[2-1])
    df_celda_semana.columns = ['C10']
    celda_no_semana = df_celda_semana.iat[0, 0]
    print(celda_no_semana)

# Si el numero de semana no coincide, ignorara los archivos 
    if celda_no_semana == week_number:
        df = pd.read_excel(os.path.join(path, file),sheet_name=0, engine='openpyxl', usecols=[0,2,17], header=None)
        # Eliminaremos las celdas vacias de las columnas con indice 0, 2 y 17
        df = df.dropna()
        # Renombramos las columnas seleccionadas
        df = df.rename(columns={0: 'CODIGO', 2: 'NOMBRE', 17: 'SALARIO'})
        # La columna de codigo la pasaremos a numeros enteros y se ignoraran los strings y otros valores
        df = df[pd.to_numeric(df['CODIGO'], errors='coerce').notnull()]
        # Para definir el nombre de la clave de obra, tomaremos el nombre del archivo hasta la aparicion del primer espacio (" ")
        df['CLAVE_OBRA'] = os.path.basename(file).split(' ')[0]
        # En cada iteracion de los archivos de excel, se crea un dataframe nuevo que se agregara a la lista de dataframes (df_list)
        df_list.append(df)

# combina todos los dataframes individuales en un único dataframe
df_final = pd.concat(df_list)

# ordenar los datos con el metodo sort
df_final = df_final.sort_values(by=['CODIGO', 'CLAVE_OBRA'],
                                    axis=0,
                                    ascending=[True,True],
                                    inplace=False)



# guarda el dataframe final en un archivo csv
df_final.to_csv(acugen_final, index=False)
print('CSV CREADO!')

# dar formato de moneda a la columna 'SALARIO'
df_final.style.format({'SALARIO': '${:,.2f}'})

# Crear un objeto de escritura de Excel
writer = pd.ExcelWriter(acugen_excel_final, engine='xlsxwriter')

# Escribir el dataframe en la primera hoja
df_final.to_excel(writer, sheet_name='SALARIO_POR_OBRA', index=False)

# Creacion de la hoja 2
df_agrupado = df_final.groupby('CODIGO')['NOMBRE','SALARIO'].sum()

# Escribir la NUEVA TABLA en la segunda hoja
df_agrupado.to_excel(writer, sheet_name='SUELDO_SEMANAL', index=True)

# Guardar el archivo
writer.save()
print('Archivo xlsx creado!')

end_time = time.time()

total_time = end_time - start_time
print("El tiempo total de ejecución fue de: {:.2f} segundos".format(total_time))


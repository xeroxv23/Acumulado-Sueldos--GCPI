# Librerias
import os
import pandas as pd
import time
import openpyxl

# Variables globales ( LAS DIRECCIONES DE LAS CARPETAS SERAN VERIFICADAS POR EL NUMERO DE SEMANA ADMINISTRATIVA, asi solo nos preocupamos por cambiar el valor de "week_number")

folder_path = '/home/xeroxv23/Documents/acumulados_sueldos_semanales/SEMANA_{}'
week_number = 3
final_path = folder_path.format(week_number)
acugen_path = '/home/xeroxv23/Documents/acumulados_sueldos_semanales/ACUGEN_SEM_{}/acumulado_sueldos.csv'
acugen_final = acugen_path.format(week_number)
acugen_excel = "/home/xeroxv23/Documents/acumulados_sueldos_semanales/ACUGEN_SEM_{}/nuevo_acugen.xlsx"
acugen_excel_final = acugen_excel.format(week_number)

# contador de tiempo
start_time = time.time()

# COMPROBACION SEMANA
df_cell = pd.read_excel('//home/xeroxv23/Documents/acumulados_sueldos_semanales/SEMANA_3/C-300 LOTE D y E OLMOS.xlsm',sheet_name=0, engine='openpyxl', header=None, skiprows=9, nrows=1, usecols=[2-1])
df_cell.columns = ['C10']
value = df_cell.iat[0, 0]


# especifica la ruta de la carpeta
path = final_path

# obtiene una lista de todos los archivos xlsm en la carpeta
files = [f for f in os.listdir(path) if f.endswith('.xlsm')]

# inicializa una lista vacía para almacenar los dataframes individuales
df_list = []

# itera a través de cada archivo xlsm y lee los datos en un dataframe
for file in files:
    df_comprobacion = pd.read_excel(os.path.join(path, file),sheet_name=0, engine='openpyxl', header=None, skiprows=9, nrows=1, usecols=[2-1])
    df_comprobacion.columns = ['C10']
    value1 = df_comprobacion.iat[0, 0]
    print(value1)

    if value1 == week_number:
        df = pd.read_excel(os.path.join(path, file),sheet_name=0, engine='openpyxl', usecols=[0,2,17], header=None)
        df = df.dropna()
        df = df.rename(columns={0: 'CODIGO', 2: 'NOMBRE', 17: 'SALARIO'})

        df = df[pd.to_numeric(df['CODIGO'], errors='coerce').notnull()]

        df['CLAVE_OBRA'] = os.path.basename(file).split(' ')[0]

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


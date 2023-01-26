import os
import pandas as pd
import time

# contador de tiempo
start_time = time.time()

# especifica la ruta de la carpeta
path = '/home/xeroxv23/Documents/acumulados_sueldos_semanales/destajos_ejemplo'

# obtiene una lista de todos los archivos xlsm en la carpeta
files = [f for f in os.listdir(path) if f.endswith('.xlsm')]

# inicializa una lista vacía para almacenar los dataframes individuales
df_list = []

# itera a través de cada archivo xlsm y lee los datos en un dataframe
for file in files:
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
df_final.to_csv('/home/xeroxv23/Documents/acumulados_sueldos_semanales/archivos_csv/acumulado_sueldos.csv', index=False)
print('CSV CREADO!')

# dar formato de moneda a la columna 'SALARIO'
df_final.style.format({'SALARIO': '${:,.2f}'})

# Crear un objeto de escritura de Excel
writer = pd.ExcelWriter("/home/xeroxv23/Documents/acumulados_sueldos_semanales/archivos_csv/nuevo_acugen.xlsx", engine='xlsxwriter')

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


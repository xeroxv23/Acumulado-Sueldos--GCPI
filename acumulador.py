import os
import pandas as pd

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
print(df_final)

# guarda el dataframe final en un archivo csv
df_final.to_csv('/home/xeroxv23/Documents/acumulados_sueldos_semanales/archivos_csv/acumulado_sueldos.csv', index=False)
print('CSV CREADO!')


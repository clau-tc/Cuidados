import sys
import os
sys.path.append('/home/clautc/Proyectos_/cuidados/adultas_mayores_y_discapacidad')
os.chdir('/home/clautc/Proyectos_/cuidados/adultas_mayores_y_discapacidad')
import xlsxwriter
from fun_procesamiento import *

# 1) Obtención data

archivo = 'tpALC_jdgo'
name_sheet = 'enf_no_transmisibles'
with open('keys/keys.txt') as k:
    keys = k.readline()
data = obtener_data_google_sheet(name_archivo=archivo, name_sheet=name_sheet, keys=keys)
# 2) procesar data
data.columns = name_columns_normal(data.columns)
data.loc[data['sex'].str.match('^B'), 'sex'] = 'Ambos sexos'
data['value'] =

if not data.sex.str.match('Ambos sexos').any():
    print('No se realizó el remplazo de categorías')
data.rename(columns={'location_name': 'pais', 'sex': 'sexo', 'age_group': 'grupo_edad'}, inplace=True)


#%%

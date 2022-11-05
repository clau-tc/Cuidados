import sys
import os

import pandas as pd

sys.path.append('/home/clautc/Proyectos_/cuidados/adultas_mayores_y_discapacidad')
os.chdir('/home/clautc/Proyectos_/cuidados/adultas_mayores_y_discapacidad')
import xlsxwriter
from fun_procesamiento import *
#%%
# 1) Obtención data
archivo = 'tpALC_jdgo'
name_sheet = 'enf_no_transmisibles'
with open('keys/keys.txt') as k:
    keys = k.readline()
data = obtener_data_google_sheet(name_archivo=archivo, name_sheet=name_sheet, keys=keys)
#%%
# 2) procesar data
#nombre columnas
data.columns = name_columns_normal(data.columns)
# remplazo de valores
data.loc[data['sex'].str.match('^B'), 'sex'] = 'Ambos sexos'
if not data.sex.str.match('Ambos sexos').any():
    print('No se realizó el remplazo de categorías')
lista = ['value', 'value_up', 'value_low']
for l in lista:
    data[l] = data[l].str.replace('\.', '', regex=True).str.replace(',', '.')
    data[l] = data[l].astype('float')
#nombre de columnas remplazo de valores
data_rename = data.rename(columns={'location_name': 'pais', 'sex': 'sexo', 'age_group': 'grupo_edad',
                     'value': 'valor', 'value_up': 'valor_max', 'value_low': 'valor_min',
                     'measure_name_en': 'medida'})
# subconjunto a usar (selección var y filtrado)
columnas_sel = ['iso3', 'sexo', 'grupo_edad', 'medida', 'valor', 'valor_min', 'valor_max', 'pais']
data_col = data_rename.loc[:, columnas_sel]
data_fil = data_col.loc[(~data_col.medida.str.match('^De') & ~data_col.sexo.str.match('^A'))]
# orden alfabético
data_fil = data_fil.sort_values(['iso3', 'grupo_edad']).reset_index(drop=True)

#%%
# 3) acceder a xlsxwriter
writer = pd.ExcelWriter('data/ent_result.xlsx', engine='xlsxwriter')
# data_fil.to_excel(writer, sheet_name='data')
wb = writer.book
# ws = writer.sheets['data']
# formato
bold = wb.add_format({'bold': 1})
# hojas para gráficos
# edades = wb.add_worksheet('edades')
combinados = wb.add_worksheet('combinados')
# crear posiciones de gráficos
posiciones = []
excel_pos = 'A1 J1 A13 J13 A26 J26 S1 AB1 S13 AB13 S26 AB26 A39 S39 A51 S51 A64 S64'.split()
for a in data_fil.medida.unique():
    for b in data_fil.grupo_edad.unique():
        posiciones.append('{} {}'.format(a, b))
dict_pos = dict(zip(posiciones, excel_pos))
# hojas para datas y construcción gráficos
medidas_lista = data_fil.medida.unique()
edad_lista = data_fil.grupo_edad.unique()
n = 0
data_medida = data_fil.groupby('medida')
for m in medidas_lista:
    m_data = data_medida.get_group(m)
    m_data = m_data.sort_values(['iso3', 'grupo_edad']).reset_index(drop=True)
    m_data_gr = m_data.groupby('grupo_edad')
    for e in edad_lista:
        e_data = m_data_gr.get_group(e)
        e_data = e_data.reset_index(drop=True)
        print(e_data)
        name_m = ''.join(list(filter(lambda c: c.isupper(), m)))
        name = name_m + e
        e_datap = e_data.pivot_table(index='iso3', columns='sexo', values='valor').reset_index(drop=False)
        e_datap.to_excel(writer, sheet_name=name)
        wsd = writer.sheets[name]
        row_max, col_max = e_datap.shape
        title = ' '.join([m, e])
        print(title)
        print(dict_pos[title])
        chart = wb.add_chart({'type': 'column'})
        chart.set_title({'name': title, 'name_font': {'size': 9}})
        chart.add_series(dict(name='Hombres', categories=[name, 1, 1, row_max, 1], values=[name, 1, 2, row_max, 2],
                              fill={'color': '#5B2C6F'}))
        chart.add_series(dict(name='Mujeres', categories=[name, 1, 1, row_max, 1], values=[name, 1, 3, row_max, 3],
                              fill={'color': '#28B463'}))
        chart.set_x_axis({'name': 'Países (ISO) y sexo',
                                'name_font': {'size': 8},
                                'num_font': {'size': 8},

                         'major_gridlines': {
                             'visible': True,
                             'line': {'width': 1.0, 'dash_type': 'dash'}
                         }}
                         )
        chart.set_y_axis({'name': 'Tasa por cada 100 mil personas',
                  'name_font': {'size': 8},
                  'num_font': {'size': 8},
                  # 'minor_unit': 0,
                  # 'major_unit': 100,
                  # 'interval_unit': 10,
                  'visible': True})
        chart.set_style(17)
        chart.set_size({'width': 568, 'height': 200})
        chart.set_legend({'position': 'bottom'})
        combinados.insert_chart(dict_pos[title], chart)
wb.close()

# pbpythn.com/pandas-pivot-report.html





# ambossexos = wb.add_worksheet('ambos_sexos')
# mujeres = wb.add_worksheet('mujeres')
# hombres = wb.add_worksheet('hombres')



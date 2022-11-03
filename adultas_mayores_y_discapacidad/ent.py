import sys
import os
sys.path.append('/home/clautc/Proyectos_/cuidados/adultas_mayores_y_discapacidad')
os.chdir('/home/clautc/Proyectos_/cuidados/adultas_mayores_y_discapacidad')
import xlsxwriter
from fun_procesamiento import *

# 1) Obtención data

archivo = 'tpALC_jdgo'
name_sheet = 'enf_no_transmisibles'
with open('/adultas_mayores_y_discapacidad/keys/keys.txt') as k:
    keys = k.readline()
data = obtener_data_google_sheet(name_archivo=archivo, name_sheet=name_sheet, keys=keys)


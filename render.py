from cmath import exp
from docxtpl import DocxTemplate
import numbers_to_letter
import pandas as pd 
import os

data = pd.read_csv('datacsv/12500_CON_ACCESORIOS.csv', encoding='utf-8')

def getDateText(date_format_num):
    months = [
        '',
        'Enero', 
        'Febrero',
        'Marzo',
        'Abril',
        'Mayo',
        'Junio',
        'Julio',
        'Agosto',
        'Septiembre',
        'Octubre',
        'Noviembre',
        'Diciembre']
    date_credit_num = date_format_num.split('/')
    return date_credit_num[0] + ' de ' + months[int(date_credit_num[1])] + ' de ' + date_credit_num[2]

for dt in data.index:

    CONTRACT_PRINCEPS       =   DocxTemplate("layouts/PROPUESTA_CRA_PRINCEPS_VFINAL3.docx")

    context_1 = {
        'clausula_1'    : False
    }
    
    # context = {

    #     'NOMBRE_COMPLETO'   :   data['NOMBRE'][dt],
    #     'REFERENCIA_BAN'    :   data['REFERENCIA'][dt],
    #     'DOMICILIO'         :   data['DOMICILIO'][dt],

    #     'VIN'               :   data['VIN'][dt],
    #     'MOTOR'             :   data['MOTOR'][dt],
    #     'MARCA'             :   data['MARCA'][dt],
    #     'MODELO'            :   data['MODELO'][dt],
    #     'COLOR'             :   data['COLOR'][dt],
        
    #     'ADEUDO'            :   data['ADEUDO'][dt],
    #     'FECHA_PAGARE'      :   data['FECHA PAGARE'][dt],
    #     'FECHA_VIGENCIA'    :   data['FECHA VIGENCIA'][dt],
    #     'FECHA_FIRMA'       :   data['FECHA FIRMA'][dt],

    #             'FECHA_CREDITO_PV'  :   getDateText(str(data['FECHA PV'][dt])),
    #     'CREDITO_PV'  :      data['CREDITO PV'],
    #     'MONTO_CREDITO_PV_NUM'  : data['MONTO PV'],
    #     'MONTO_CREDITO_PV_LETRA'  :'',
        
    #     'FECHA_CREDITO_FS'  :   getDateText(str(data['FECHA FS'][dt])),
    #     'CREDITO_FS'        :   data['CREDITO FS'][dt],
    #     'MONTO_FS_NUM'      :   data['MONTO FS'][dt],
    #     'MONTO_FS_LETRA'    :   "({} PESOS {}/100 M.N.)".format(
    #         numbers_to_letter.numero_a_letras(int(float(data['MONTO FS'][dt]))).upper(),
    #         str(data['MONTO FS'][dt]).split('.')[1]),
        
    #     'FECHA_CREDITO_GPS'  : getDateText(str(data['FECHA GPS'][dt])),
    #     'CREDITO_GPS'  :        data['CREDITO GPS'],
    #     'MONTO_CREDITO_GPS_NUM'  : data['MONTO GPS'],
    #     'MONTO_CREDITO_GPS_LETRA'  : "({} PESOS {}/100 M.N.)".format(
    #         numbers_to_letter.numero_a_letras(int(float(data['MONTO GPS'][dt]))).upper(),
    #         str(data['MONTO GPS'][dt]).split('.')[1]),
        
    #     'FECHA_CREDITO_GASTOS'  : getDateText(str(data['FECHA GASTOS'][dt])),
    #     'CREDITO_GASTOS'  :   data['CREDITO GASTO'],
    #     'MONTO_CREDITO_GASTOS_NUM'  : data['MONTO GASTO'],
    #     'MONTO_CREDITO_GASTOS_LETRA'  : "({} PESOS {}/100 M.N.)".format(
    #         numbers_to_letter.numero_a_letras(int(float(data['MONTO GASTO'][dt]))).upper(),
    #         str(data['MONTO GASTO'][dt]).split('.')[1]),
        
    #     'FECHA_CREDITO_R2021'  : getDateText(str(data['FECHA R2021'][dt])),
    #     'CREDITO_R2021'  :  data['CREDITO R2021'],
    #     'MONTO_CREDITO_R2021_NUM'  : data['MONTO R2021'],
    #     'MONTO_CREDITO_R2021_LETRA'  : "({} PESOS {}/100 M.N.)".format(
    #         numbers_to_letter.numero_a_letras(int(float(data['MONTO R2021'][dt]))).upper(),
    #         str(data['MONTO R2021'][dt]).split('.')[1]),
        
    #     'FECHA_CREDITO_ENRUTA'  : getDateText(str(data['FECHA ENRUTA'][dt])),
    #     'CREDITO_ENRUTA'  :  data['CREDITO ENRUTA'],
    #     'MONTO_CREDITO_ENRUTA_NUM'  :data['MONTO ENRUTA'],
    #     'MONTO_CREDITO_ENRUTA_LETRA'  : "({} PESOS {}/100 M.N.)".format(
    #         numbers_to_letter.numero_a_letras(int(float(data['MONTO ENRUTA'][dt]))).upper(),
    #         str(data['MONTO ENRUTA'][dt]).split('.')[1]),

    # }


    fileDir = 'contratosPRUEBA/'

    #print(nombreRuta)
    try:
        os.stat('contratosPRUEBA/')
    except:
        os.mkdir('contratosPRUEBA/')
    
    try:
        os.stat(fileDir)
    except:
        os.mkdir(fileDir)

    CONTRACT_PRINCEPS.render(context_1)
    CONTRACT_PRINCEPS.save(fileDir + '/' + 'PRUEBA PARRAFOS_' + str(data['NOMBRE'][dt]) + "_" + str(data['CREDITO FS'][dt]) + ".docx")

    break


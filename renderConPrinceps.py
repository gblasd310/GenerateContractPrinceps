from multiprocessing import context
from docxtpl import DocxTemplate
import numbers_to_letter
import pandas as pd 
import os

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
    #print(date_credit_num)
    return date_credit_num[0] + ' de ' + months[int(date_credit_num[1])] + ' de ' + date_credit_num[2] + ','

data = pd.read_csv('datacsv/12500_CON_ACCESORIOS.csv', encoding='utf-8')

#print(data)

for dt in data.index:

    CONTRACT_PRINCEPS       =   DocxTemplate("layouts/PROPUESTA_CRA_PRINCEPS_VFINAL3.docx")

    context = {
        'NOMBRE_COMPLETO'   :   data['NOMBRE'][dt],
        'REFERENCIA_DV'     :   str(data['REFERENCIA MAS DV'][dt]).zfill(11),
        'DOMICILIO'         :   data['DOMICILIO'][dt],

        'VIN'               :   data['VIN'][dt],
        'MOTOR'             :   data['MOTOR'][dt],
        'MARCA'             :   data['MARCA'][dt],
        'MODELO'            :   data['MODELO'][dt],
        'COLOR'             :   data['COLOR'][dt],
        
        'ADEUDO'            :   data['ADEUDO'][dt],
        'FECHA_PAGARE'      :   str(getDateText(str(data['FECHA PAGARE'][dt]))),
        'FECHA_VIGENCIA'    :   str(getDateText(str(data['FECHA VIGENCIA'][dt]))),
        'FECHA_FIRMA'       :   str(getDateText(str(data['FECHA FIRMA'][dt]))),
        'clausula_LC'       :   False,
        # 'clausula_PV'       :   False,
        # 'clausula_FS'       :   False,
        # 'clausula_GPS'      :   False,
        # 'clausula_GASTOS'   :   False,
        # 'clausula_R2021'    :   False,
        # 'clausula_ENRUTA'   :   False

    }

    #print(str(data['CREDITO PV'][dt]))
    if str(data['CREDITO PV'][dt]) != '0':
        context['clausula_PV']  = True
        context['FECHA_CREDITO_PV']  =   str(getDateText(str(data['FECHA PV'][dt])))
        context['CREDITO_PV']  =     str(data['CREDITO PV'][dt]) + ','
        context['MONTO_PV_NUM']  = data['MONTO PV'][dt]
        context['MONTO_PV_LETRA']  = "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO PV'][dt]))).upper(),
            str(data['MONTO PV'][dt]).split('.')[1])

    if str(data['CREDITO FS'][dt]) != '0':
        context['clausula_FS']  = True
        context['FECHA_CREDITO_FS']  =   getDateText(str(data['FECHA FS'][dt]))
        context['CREDITO_FS']        =   str(data['CREDITO FS'][dt]) + ','
        context['MONTO_FS_NUM' ]     =   data['MONTO FS'][dt]
        context['MONTO_FS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO FS'][dt]))).upper(),
            str(data['MONTO FS'][dt]).split('.')[1])

    if str(data['CREDITO GPS'][dt]) != '0':
        context['clausula_GPS']  = True
        context['FECHA_CREDITO_GPS']  =   getDateText(str(data['FECHA GPS'][dt]))
        context['CREDITO_GPS']        =   str(data['CREDITO GPS'][dt]) + ','
        context['MONTO_GPS_NUM' ]     =   data['MONTO GPS'][dt]
        context['MONTO_GPS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO GPS'][dt]))).upper(),
            str(data['MONTO GPS'][dt]).split('.')[1])

    if str(data['CREDITO GASTOS'][dt]) != '0':
        context['clausula_GASTOS']  = True
        context['FECHA_CREDITO_GASTOS']  =   getDateText(str(data['FECHA GASTOS'][dt]))
        context['CREDITO_GASTOS']        =   str(data['CREDITO GASTOS'][dt]) + ','
        context['MONTO_GASTOS_NUM' ]     =   data['MONTO GASTOS'][dt]
        context['MONTO_GASTOS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO GASTOS'][dt]))).upper(),
            str(data['MONTO GASTOS'][dt]).split('.')[1])

    if str(data['CREDITO R2021'][dt]) != '0':
        context['clausula_R2021']  = True
        context['FECHA_CREDITO_R2021']  =   getDateText(str(data['FECHA 2021'][dt]))
        context['CREDITO_R2021']        =   str(data['CREDITO R2021'][dt]) + ','
        context['MONTO_R2021_NUM' ]     =   data['MONTO R2021'][dt]
        context['MONTO_R2021_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO R2021'][dt]))).upper(),
            str(data['MONTO R2021'][dt]).split('.')[1])
    
    if str(data['CREDITO ENRUTA'][dt]) != '0':
        context['clausula_ENRUTA']  = True
        context['FECHA_CREDITO_ENRUTA']  =   getDateText(str(data['FECHA ENRUTA'][dt]))
        context['CREDITO_ENRUTA']        =   str(data['CREDITO ENRUTA'][dt]) + ','
        context['MONTO_ENRUTA_NUM' ]     =   data['MONTO ENRUTA'][dt]
        context['MONTO_ENRUTA_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO ENRUTA'][dt]))).upper(),
            str(data['MONTO ENRUTA'][dt]).split('.')[1])



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

    print(context)

    CONTRACT_PRINCEPS.render(context)
    CONTRACT_PRINCEPS.save(fileDir + '/' + 'PRINCEPS_' + str(data['NOMBRE'][dt]) + "_" + str(data['CREDITO'][dt]) + ".docx")
    print('PRINCEPS_' + str(data['NOMBRE'][dt]) + "_" + str(data['CREDITO'][dt]) + ".docx")

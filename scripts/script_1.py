from enum import Flag
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
    return date_credit_num[0] + ' de ' + months[int(date_credit_num[1])] + ' de ' + date_credit_num[2]

data = pd.read_csv('datacsv/12500_CON_ACCESORIOS.csv', encoding='utf-8')

#print(data)


for dt in data.index:

    CONTRACT_PRINCEPS       =   DocxTemplate("layouts/PROPUESTA_CRA_PRINCEPS_VFINAL3 - LAYOUT - CARTACONDONACION.docx")

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
        'FECHA_FIRMA'       :   "01 de Octubre de 2022", #str(getDateText(str(data['FECHA FIRMA'][dt]))),

        
        'clausulaDOSENGPV'   :   False,

        'CREDITO_ANTERIOR'  :   data['CREDITO ANTERIOR'][dt],
        'FECHA_CTO_ANTERIOR':   str(getDateText(str(data['FECHA CREDITO ANTERIOR'][dt]))),
        'MONTO_CTO_ANTERIOR_NUM':   '{:,.2f}'.format(float(data['MONTO CREDITO ANTERIOR'][dt])),
        'MONTO_CTO_ANTERIOR_LETRA':  "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO CREDITO ANTERIOR'][dt]))).upper(),
            str(data['MONTO CREDITO ANTERIOR'][dt]).split('.')[1])

    }

    if  str(data['MONTO CONDONACION'][dt]) != "NO APLICA":
        context['carta_condonacion'] = True
        context['CREDITO']  = data['CREDITO'][dt]
        context['NMENSUALIDADES'] = data['TOTAL PAGOS'][dt]
        context['MONTOMENSUALIDAD'] = "${:,.2f} ({} PESOS {}/100 M.N.)".format(
            float(data['MONTO MENSUALIDAD'][dt]),
            numbers_to_letter.numero_a_letras(int(float(data['MONTO MENSUALIDAD'][dt]))).upper(),
            str(data['MONTO MENSUALIDAD'][dt]).split('.')[1])
        context['MONTOCONDONACION'] = "${:,.2f} ({} PESOS {}/100 M.N.)".format(
            float(data['MONTO CONDONACION'][dt]),
            numbers_to_letter.numero_a_letras(int(float(data['MONTO CONDONACION'][dt]))).upper(),
            str(data['MONTO CONDONACION'][dt]).split('.')[1])

    #ENGANCHE PV ARRENDAMIENTO
    if str(data['CREDITO PV2'][dt]) != '0':
        context['clausulaPV2']  = True
        context['FECHA_CREDITO_PV2']  =   str(getDateText(str(data['FECHA PV2'][dt])))
        context['CREDITO_PV2']  =     str(data['CREDITO PV2'][dt]) + ','
        context['MONTO_PV_NUM']  = '{:,.2f}'.format(float(data['MONTO PV2'][dt]))
        context['MONTO_PV_LETRA']  = "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO PV2'][dt]))).upper(),
            str(data['MONTO PV2'][dt]).split('.')[1])

    # OPCION 3 CREDITO DE ADQUISICION
    if str(data['CREDITO PV3'][dt]) != '0':
        context['clausulaPV3']  = True
        context['FECHA_CREDITO_PV3']  =   getDateText(str(data['FECHA PV3'][dt]))
        context['CREDITO_PV3']        =   str(data['CREDITO PV3'][dt]) + ','
        context['MONTO_FS_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO FS'][dt]))
        context['MONTO_FS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO PV3'][dt]))).upper(),
            str(data['MONTO PV3'][dt]).split('.')[1])

    if str(data['CREDITO ENG PV'][dt]) != '0':
        context['clausulaDOSENGPV']  = True
        context['FECHA_CREDITO_ENG_PV']  =   getDateText(str(data['FECHA ENG PV'][dt]))
        context['CREDITO_ENG_PV']        =   str(data['CREDITO ENG PV'][dt]) + ','
        context['MONTO_ENG_PV_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO ENG PV'][dt]))
        context['MONTO_ENG_PV_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO ENG PV'][dt]))).upper(),
            str(data['MONTO ENG PV'][dt]).split('.')[1])
    
    if str(data['CREDITO FS'][dt]) != '0':
        context['clausulaTRESFS']  = True
        context['FECHA_CREDITO_FS']  =   getDateText(str(data['FECHA FS'][dt]))
        context['CREDITO_FS']        =   str(data['CREDITO FS'][dt]) + ','
        context['MONTO_ENG_FS_NUM' ]     =  '{:,.2f}'.format(float( data['MONTO FS'][dt]))
        context['MONTO_ENG_FS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO FS'][dt]))).upper(),
            str(data['MONTO FS'][dt]).split('.')[1])

    if str(data['CREDITO GPS'][dt]) != '0':
        context['clausulaCUATROGPS']  = True
        context['FECHA_CREDITO_GPS']  =   getDateText(str(data['FECHA GPS'][dt]))
        context['CREDITO_GPS']        =   str(data['CREDITO GPS'][dt]) + ','
        context['MONTO_GPS_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO GPS'][dt]))
        context['MONTO_GPS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO GPS'][dt]))).upper(),
            str(data['MONTO GPS'][dt]).split('.')[1])

    if str(data['CREDITO GASTOS'][dt]) != '0':
        context['clausulaCINCOGASTOS']  = True
        context['FECHA_CREDITO_GASTOS']  =   getDateText(str(data['FECHA GASTOS'][dt]))
        context['CREDITO_GASTOS']        =   str(data['CREDITO GASTOS'][dt]) + ','
        context['MONTO_GASTOS_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO GASTOS'][dt]))
        context['MONTO_GASTOS_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO GASTOS'][dt]))).upper(),
            str(data['MONTO GASTOS'][dt]).split('.')[1])

    if str(data['CREDITO CESION PV'][dt]) != '0':
        context['clausulaSEISCESIONPV']  = True
        context['FECHA_CESION_PV']  =   getDateText(str(data['FECHA CESION PV'][dt]))
        context['CREDITO_CESION_PV']        =   str(data['CREDITO CESION PV'][dt]) + ','
        context['MONTO_CESION_PV_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO CESION PV'][dt]))
        context['MONTO_CESION_PV_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO CESION PV'][dt]))).upper(),
            str(data['MONTO CESION PV'][dt]).split('.')[1])
    
    if str(data['CREDITO RENOVACION 2021'][dt]) != '0':
        context['clausulaSIETER2021']  = True
        context['FECHA_CREDITO_R2021']  =   getDateText(str(data['FECHA RENOVACION 2021'][dt]))
        context['CREDITO_R2021']        =   str(data['CREDITO RENOVACION 2021'][dt]) + ','
        context['MONTO_R2021_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO RENOVACION 2021'][dt]))
        context['MONTO_R2021_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO RENOVACION 2021'][dt]))).upper(),
            str(data['MONTO RENOVACION 2021'][dt]).split('.')[1])
    
    if str(data['CREDITO ENRUTA'][dt]) != '0':
        context['clausulaOCHOENRUTA']  = True
        context['FECHA_CREDITO_ENRUTA']  =   getDateText(str(data['FECHA ENRUTA'][dt]))
        context['CREDITO_ENRUTA']        =   str(data['CREDITO ENRUTA'][dt]) + ','
        context['MONTO_ENRUTA_NUM' ]     =   '{:,.2f}'.format(float(data['MONTO ENRUTA'][dt]))
        context['MONTO_ENRUTA_LETRA']    =   "({} PESOS {}/100 M.N.)".format(
            numbers_to_letter.numero_a_letras(int(float(data['MONTO ENRUTA'][dt]))).upper(),
            str(data['MONTO ENRUTA'][dt]).split('.')[1])

    if context['clausulaDOSENGPV']:
        context['clausulaSEISCESIONPV '] = False
        


    fileDir = 'C:/Users/Gustavo Blas/OneDrive - Financera Sustentable de México SA de CV SFP/CONTRATOS OCTUBRE/contratosPRINCEPS_12500_CON_ACCESORIOS_V2/'

    #print(nombreRuta)
    try:
        os.stat('C:/Users/Gustavo Blas/OneDrive - Financera Sustentable de México SA de CV SFP/CONTRATOS OCTUBRE/contratosPRINCEPS_12500_CON_ACCESORIOS_V2/')
    except:
        os.mkdir('C:/Users/Gustavo Blas/OneDrive - Financera Sustentable de México SA de CV SFP/CONTRATOS OCTUBRE/contratosPRINCEPS_12500_CON_ACCESORIOS_V2/')
    
    try:
        os.stat(fileDir)
    except:
        os.mkdir(fileDir)

    #print(context)

    CONTRACT_PRINCEPS.render(context)
    CONTRACT_PRINCEPS.save(fileDir + '/' + 'PRINCEPS_' + str(data['NOMBRE'][dt]) + "_" + str(data['CREDITO'][dt]) + ".docx")
    print('PRINCEPS_' + str(data['NOMBRE'][dt]) + "_" + str(data['CREDITO'][dt]) + ".docx")

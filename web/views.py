from django.shortcuts import render
from .i18n import *

from .models import *
# importar el modelo de tabla User 
from django.contrib.auth.models import User,Group
# importar librerias que permitan la validacion del login
from django.contrib.auth import authenticate,logout,login as login_aut
# importar libreria decoradora que permite evitar el ingreso de usuarios a las paginas web
from django.contrib.auth.decorators import login_required, permission_required

from django.shortcuts import redirect
from datetime import datetime, timedelta
import uuid
import math

from openpyxl import Workbook,load_workbook
from django.http import HttpResponse
from io import BytesIO

import qrcode
from io import BytesIO
from django.core.files import File
from PIL import Image,ImageDraw
from django.core.mail import EmailMessage
from django.http import HttpResponse
from django.core.files.base import ContentFile
import pandas as pd
# Create your views here.
documento_inf=[]
monto_buscar_inf=[]
diferencia_inf=[]
fecha_inf=[]
tc_inf=[]

def inicio(request):
    request.session["datos"]=""
    x={}

    # x["valor"]=request.session["datos"]
    # x["habitacion"]=habi(0)
    comentarios=Comentario.objects.all()
    mensaje={'comentarios':comentarios}
    return render(request,"index.html",mensaje)

def procesamiento(request):
    print("entro")
    contexto={}
    if request.method == 'POST':
        archivo1 = request.FILES.get('archivo1')
        archivo2 = request.FILES.get('archivo2')
        archivo3 = request.FILES.get('archivo3')
        archivo4 = request.FILES.get('archivo4') # visa dolar
        archivo5 = request.FILES.get('archivo5')
        archivo7 = request.FILES.get('archivo7')
        if archivo1:
            print("ES VALIDO")
            x=pd.read_excel(archivo1,usecols=[1,2,3,4,5,6,7,8,9,10], nrows=1500)
            diccionario = {}
            i=0
            for fila in x.index:
                i+= 1
                texto = f"{i} - {x.at[fila, 'Unnamed: 3']} - {x.at[fila, 'Unnamed: 5']} - {x.at[fila,'Unnamed: 8']}"
                if "US$" in texto:
                    print(texto)
                    codigo_buscar = x.at[fila, 'Unnamed: 3']
                    valor_buscado = abs(int(x.at[fila, 'Unnamed: 8']))
                    fecha_erp = x.at[fila, 'Unnamed: 10']
                    #print(f"CODIGO A BUSCAR: {codigo_buscar} - Valor Buscado:{valor_buscado} ")
                    print("-----------------------------------")
                    if "Amex US$" in texto:
                        cantidad=0
                        suma = 0
                        if archivo2:
                            print("Archivo de AMEX US$")
                            amex=pd.read_excel(archivo2,usecols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14], nrows=2500)
                            i_amex=0
                            print("Valores Archivo AMEX US$ -----------------------------------")
                            sw=0
                            for fila_amex in amex.index:
                                codigo = amex.at[fila_amex, 'Unnamed: 2']
                                i_amex = i_amex +1
                                texto_amex = f"{i_amex} - {amex.at[fila_amex, 'Unnamed: 2']} - {amex.at[fila_amex, 'Unnamed: 12']} - {amex.at[fila_amex,'Unnamed: 13']} - {amex.at[fila_amex,'Unnamed: 14']}"
                                if codigo_buscar==codigo:
                                    cantidad = cantidad +1
                                    print("----> Encontrada la Cuenta:")
                                    if not math.isnan(amex.at[fila_amex, 'Unnamed: 12']): 
                                        if valor_buscado==abs(int(amex.at[fila_amex, 'Unnamed: 12'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en OTRA MONEDA {amex.at[fila_amex, 'Unnamed: 12']}")
                                            suma = suma + abs(int(amex.at[fila_amex, 'Unnamed: 12']))
                                    if not math.isnan(amex.at[fila_amex, 'Unnamed: 13']):                                     
                                        if valor_buscado==abs(int(amex.at[fila_amex, 'Unnamed: 13'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO {amex.at[fila_amex, 'Unnamed: 13']}")
                                            suma = suma + abs(int(amex.at[fila_amex, 'Unnamed: 13']))
                                    if not math.isnan(amex.at[fila_amex, 'Unnamed: 14']): 
                                        if valor_buscado==abs(int(amex.at[fila_amex, 'Unnamed: 14'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO CORREGIDO {amex.at[fila_amex, 'Unnamed: 14']}")
                                            suma = suma + abs(int(amex.at[fila_amex, 'Unnamed: 14']))                                    
                                    print(texto_amex)
                                    sw=1
                                    diccionario[f"{i}"]={"tarjeta":"Amex US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"encontrado"}
                                    continue
                            if(sw==0):
                                print(f"NO ENCONTRO EL CODIGO : {codigo_buscar} CUYO VALOR ES DE {valor_buscado}")
                                tc_inf.append("Amex US$")
                                documento_inf.append(abs(int(codigo_buscar)))
                                monto_buscar_inf.append(valor_buscado)
                                fecha_inf.append(fecha_erp)
                                diccionario[f"{i}"]={"tarjeta":"Amex US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"no encontrado"}                                               
                    if "Dinners US$" in texto:
                        cantidad=0
                        suma = 0
                        if archivo3:
                            print("Archivo de Dinners US$")
                            dinners=pd.read_excel(archivo3,usecols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14], nrows=2500)
                            i_dinners=0
                            print("Valores Archivo DINNERS US$ -----------------------------------")
                            sw=0
                            for fila_dinners in dinners.index:
                                codigo = dinners.at[fila_dinners, 'Unnamed: 2']
                                i_dinners = i_dinners +1
                                texto_dinners = f"{i_dinners} - {dinners.at[fila_dinners, 'Unnamed: 2']} - {dinners.at[fila_dinners, 'Unnamed: 12']} - {dinners.at[fila_dinners,'Unnamed: 13']} - {dinners.at[fila_dinners,'Unnamed: 14']}"
                                if codigo_buscar==codigo:
                                    cantidad = cantidad +1
                                    print("----> Encontrada la Cuenta:")
                                    if not math.isnan(dinners.at[fila_dinners, 'Unnamed: 12']): 
                                        if valor_buscado==abs(int(dinners.at[fila_dinners, 'Unnamed: 12'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en OTRA MONEDA {dinners.at[fila_dinners, 'Unnamed: 12']}")
                                            suma = suma + abs(int(dinners.at[fila_dinners, 'Unnamed: 12']))                                            
                                            
                                    if not math.isnan(dinners.at[fila_dinners, 'Unnamed: 13']):                                     
                                        if valor_buscado==abs(int(dinners.at[fila_dinners, 'Unnamed: 13'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO {dinners.at[fila_dinners, 'Unnamed: 13']}")
                                            suma = suma + abs(int(dinners.at[fila_dinners, 'Unnamed: 13']))
                                            
                                    if not math.isnan(dinners.at[fila_dinners, 'Unnamed: 14']): 
                                        if valor_buscado==abs(int(dinners.at[fila_dinners, 'Unnamed: 14'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO CORREGIDO {dinners.at[fila_dinners, 'Unnamed: 14']}")
                                            suma = suma + abs(int(dinners.at[fila_dinners, 'Unnamed: 14']))                                    
                                    print(texto_dinners)
                                    sw=1
                                    diccionario[f"{i}"]={"tarjeta":"Dinners US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"encontrado"}
                            if(sw==0):
                                print(f"NO ENCONTRO EL CODIGO : {codigo_buscar} CUYO VALOR ES DE {valor_buscado}")
                                tc_inf.append("Dinners US$")
                                documento_inf.append(abs(int(codigo_buscar)))
                                monto_buscar_inf.append(valor_buscado)
                                fecha_inf.append(fecha_erp)
                                diccionario[f"{i}"]={"tarjeta":"Dinners US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"no encontrado"}                                             
                    if "Master Card US$" in texto:
                        cantidad=0
                        suma = 0
                        if archivo4:
                            print("Archivo de Master Card US$")
                            mc=pd.read_excel(archivo5,usecols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14], nrows=2500)
                            i_mc=0
                            print("Valores Archivo Master Card US$ -----------------------------------")
                            sw=0
                            for fila_mc in mc.index:
                                codigo = mc.at[fila_mc, 'Unnamed: 2']
                                i_mc = i_mc +1
                                texto_mc = f"{i_mc} - {mc.at[fila_mc, 'Unnamed: 2']} - {mc.at[fila_mc, 'Unnamed: 12']} - {mc.at[fila_mc,'Unnamed: 13']} - {mc.at[fila_mc,'Unnamed: 14']}"
                                if codigo_buscar==codigo:
                                    cantidad = cantidad +1
                                    print("----> Encontrada la Cuenta:")
                                    if not math.isnan(mc.at[fila_mc, 'Unnamed: 12']): 
                                        if valor_buscado==abs(int(mc.at[fila_mc, 'Unnamed: 12'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en OTRA MONEDA {mc.at[fila_mc, 'Unnamed: 12']}")
                                            suma = suma + abs(int(mc.at[fila_mc, 'Unnamed: 12']))                                            
                                            
                                    if not math.isnan(mc.at[fila_mc, 'Unnamed: 13']):                                     
                                        if valor_buscado==abs(int(mc.at[fila_mc, 'Unnamed: 13'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO {mc.at[fila_mc, 'Unnamed: 13']}")
                                            suma = suma + abs(int(mc.at[fila_mc, 'Unnamed: 13']))
                                            
                                    if not math.isnan(mc.at[fila_mc, 'Unnamed: 14']): 
                                        if valor_buscado==abs(int(mc.at[fila_mc, 'Unnamed: 14'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO CORREGIDO {mc.at[fila_mc, 'Unnamed: 14']}")
                                            suma = suma + abs(int(mc.at[fila_mc, 'Unnamed: 14']))                                    
                                    print(texto_mc)
                                    sw=1
                                    diccionario[f"{i}"]={"tarjeta":"Master Card US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"encontrado"}
                            if(sw==0):
                                print(f"NO ENCONTRO EL CODIGO : {codigo_buscar} CUYO VALOR ES DE {valor_buscado}")
                                tc_inf.append("Master Card US$")
                                documento_inf.append(abs(int(codigo_buscar)))
                                monto_buscar_inf.append(valor_buscado)
                                fecha_inf.append(fecha_erp)
                                diccionario[f"{i}"]={"tarjeta":"Master Card US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"no encontrado"}                                             

                    if "Visa US$" in texto:
                        cantidad=0
                        suma = 0
                        if archivo4:
                            print("Archivo de Visa US$")
                            mc=pd.read_excel(archivo4,usecols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14], nrows=2500)
                            i_mc=0
                            print("Valores Archivo Visa US$ -----------------------------------")
                            sw=0
                            for fila_mc in mc.index:
                                codigo = mc.at[fila_mc, 'Unnamed: 2']
                                i_mc = i_mc +1
                                texto_mc = f"{i_mc} - {mc.at[fila_mc, 'Unnamed: 2']} - {mc.at[fila_mc, 'Unnamed: 12']} - {mc.at[fila_mc,'Unnamed: 13']} - {mc.at[fila_mc,'Unnamed: 14']}"
                                if codigo_buscar==codigo:
                                    cantidad = cantidad +1
                                    print("----> Encontrada la Cuenta:")
                                    if not math.isnan(mc.at[fila_mc, 'Unnamed: 12']): 
                                        if valor_buscado==abs(int(mc.at[fila_mc, 'Unnamed: 12'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en OTRA MONEDA {mc.at[fila_mc, 'Unnamed: 12']}")
                                            suma = suma + abs(int(mc.at[fila_mc, 'Unnamed: 12']))                                            
                                            
                                    if not math.isnan(mc.at[fila_mc, 'Unnamed: 13']):                                     
                                        if valor_buscado==abs(int(mc.at[fila_mc, 'Unnamed: 13'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO {mc.at[fila_mc, 'Unnamed: 13']}")
                                            suma = suma + abs(int(mc.at[fila_mc, 'Unnamed: 13']))
                                            
                                    if not math.isnan(mc.at[fila_mc, 'Unnamed: 14']): 
                                        if valor_buscado==abs(int(mc.at[fila_mc, 'Unnamed: 14'])):
                                            print(f"Valor Buscado: {valor_buscado} esta presente en SALDO CORREGIDO {mc.at[fila_mc, 'Unnamed: 14']}")
                                            suma = suma + abs(int(mc.at[fila_mc, 'Unnamed: 14']))                                    
                                    print(texto_mc)
                                    sw=1
                                    diccionario[f"{i}"]={"tarjeta":"Visa US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"encontrado"}
                            if(sw==0):
                                print(f"NO ENCONTRO EL CODIGO : {codigo_buscar} CUYO VALOR ES DE {valor_buscado}")
                                tc_inf.append("Visa US$")
                                documento_inf.append(abs(int(codigo_buscar)))
                                monto_buscar_inf.append(valor_buscado)
                                fecha_inf.append(fecha_erp)
                                diccionario[f"{i}"]={"tarjeta":"Visa US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"no encontrado"}                                             

                continue
                clave = x.at[fila, 'Unnamed: 3']
                valor = x.at[fila, 'Unnamed: 6']
                    # Verifica que la clave no sea nan
                if pd.isna(clave):
                    # print("Clave nula, se omite:", clave)
                    continue
                if x.at[fila, 'Unnamed: 3'] in diccionario:
                    # print("La clave ya existe:", x.at[fila, 'Unnamed: 3'])
                    diccionario[x.at[fila, 'Unnamed: 3']] = diccionario[x.at[fila, 'Unnamed: 3']] + [x.at[fila, 'Unnamed: 6']]        
                else:
                    # print("La clave no existe, se crea:", x.at[fila, 'Unnamed: 3']) 
                    # print("Valor:", x.at[fila, 'Unnamed: 6']) 
                    # print("Tipo de valor:", type(x.at[fila, 'Unnamed: 6']))
                    if str(valor).isnumeric():
                        diccionario[str(clave)] = [valor]      
                    # diccionario[str(clave)].append(valor)
                
                print(texto)
                i += 1
                # for columna in x.columns:
                #     valor = x.at[fila, columna]
                #     print("columna", columna)
                #     print(f"Fila {fila}, Columna {columna}: {valor}")
        if archivo7:
                print("procesar Transbank con ERP (archivo7 vcon archivo1)")
                print("ES VALIDO ARCHIVO 7")
                TRANS   =pd.read_excel(archivo7,usecols=[1,2,3,4,5,6,7,8,9,10], nrows=1500) #Archivo Transbank
                diccionario_2 = {}
                i2=0
                for fila in TRANS.index:
                    i2+= 1
                    fecha = TRANS.at[fila, 'Unnamed: 2']
                    tipo_tarjeta = TRANS.at[fila, 'Unnamed: 3']
                    monto_original = TRANS.at[fila, 'Unnamed: 6']
                    codigo_aut = TRANS.at[fila, 'Unnamed: 7']
                    texto = f"{i2} - {fecha} - {tipo_tarjeta} - {monto_original} - {codigo_aut}"
                    print(f"TRANSBANK: -----> {texto}")
                    ERP=pd.read_excel(archivo1,usecols=[1,2,3,4,5,6,7,8,9,10], nrows=1500) #Archivo ERP
                    sw=0
                    for fila_erp in ERP.index:
                        monto =  ERP.at[fila_erp, 'Unnamed: 8']
                        codigo_erp = ERP.at[fila_erp, 'Unnamed: 9']
                        if not pd.isna(tipo_tarjeta):                            
                            if str(codigo_aut) in str(codigo_erp):
                                sw=1
                                print(f"Encontrado ..... (monto original: {monto_original}) - (monto ERP: {monto})")                                
                                diccionario_2[f"{i2}"]={"tarjeta":tipo_tarjeta,"codigo":str(codigo_aut),"Valor":abs(int(monto)),"Status":"encontrado"}    
                    if sw==0:
                        print("No se encontro ------")
                        diccionario_2[f"{i2}"]={"tarjeta":tipo_tarjeta,"codigo":(codigo_aut),"Valor":monto_original,"Status":"no encontrado"}
                    # if "US$" in texto:
                    #     print(texto)
                    #     codigo_buscar = x.at[fila, 'Unnamed: 3']
                    #     valor_buscado = abs(int(x.at[fila, 'Unnamed: 8']))
                    #     #print(f"CODIGO A BUSCAR: {codigo_buscar} - Valor Buscado:{valor_buscado} ")
                    #     print("-----------------------------------")
                    #     if "Amex US$" in texto:
                    #         cantidad=0
                    #         suma = 0
                    #         if archivo2:
                    #             print("Archivo de AMEX US$")
                    #             amex=pd.read_excel(archivo2,usecols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14], nrows=2500)
                    #             i_amex=0
                    #             print("Valores Archivo AMEX US$ -----------------------------------")
                    #             sw=0
                    #             for fila_amex in amex.index:
                    #                 codigo = amex.at[fila_amex, 'Unnamed: 2']
                    #                 i_amex = i_amex +1
                    #                 texto_amex = f"{i_amex} - {amex.at[fila_amex, 'Unnamed: 2']} - {amex.at[fila_amex, 'Unnamed: 12']} - {amex.at[fila_amex,'Unnamed: 13']} - {amex.at[fila_amex,'Unnamed: 14']}"
                    #                 if codigo_buscar==codigo:
                    #                     cantidad = cantidad +1
                    #                     print("----> Encontrada la Cuenta:")
                    #                     if not math.isnan(amex.at[fila_amex, 'Unnamed: 12']): 
                    #                         if valor_buscado==abs(int(amex.at[fila_amex, 'Unnamed: 12'])):
                    #                             print(f"Valor Buscado: {valor_buscado} esta presente en OTRA MONEDA {amex.at[fila_amex, 'Unnamed: 12']}")
                    #                             suma = suma + abs(int(amex.at[fila_amex, 'Unnamed: 12']))
                    #                     if not math.isnan(amex.at[fila_amex, 'Unnamed: 13']):                                     
                    #                         if valor_buscado==abs(int(amex.at[fila_amex, 'Unnamed: 13'])):
                    #                             print(f"Valor Buscado: {valor_buscado} esta presente en SALDO {amex.at[fila_amex, 'Unnamed: 13']}")
                    #                             suma = suma + abs(int(amex.at[fila_amex, 'Unnamed: 13']))
                    #                     if not math.isnan(amex.at[fila_amex, 'Unnamed: 14']): 
                    #                         if valor_buscado==abs(int(amex.at[fila_amex, 'Unnamed: 14'])):
                    #                             print(f"Valor Buscado: {valor_buscado} esta presente en SALDO CORREGIDO {amex.at[fila_amex, 'Unnamed: 14']}")
                    #                             suma = suma + abs(int(amex.at[fila_amex, 'Unnamed: 14']))                                    
                    #                     print(texto_amex)
                    #                     sw=1
                    #                     diccionario[f"{i}"]={"tarjeta":"Amex US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"encontrado"}
                    #                     continue
                    #             if(sw==0):
                    #                 print(f"NO ENCONTRO EL CODIGO : {codigo_buscar} CUYO VALOR ES DE {valor_buscado}")
                    #                 diccionario[f"{i}"]={"tarjeta":"Amex US$","codigo":abs(int(codigo_buscar)),"Valor":valor_buscado,"Status":"no encontrado"}                                               
                
        diccionario_temp=[
                {'Amex US$': 'Amex US$', 'Valor US': 0,'Valor TBK':0,'Diferencia':0},
                {'Dinners US$':'Dinners US$', 'Valor US': 0,'Valor TBK':0,'Diferencia':0},
                {'Master Card US$': 'Master Card US$', 'Valor US': 0,'Valor TBK':0,'Diferencia':0},
                {'Visa US$': 'Visa US$', 'Valor US': 0,'Valor TBK':0,'Diferencia':0},
                {'Total': 'Total', 'Valor US': 0,'Valor TBK':0,'Diferencia':0}
            ]
        total_amex=0
        total_dinners=0
        total_mc=0
        total_visa=0
        for clave, valor in diccionario.items():
            print(f"Clave: {clave} -> Valor: {valor} -> Dato Valor 1:{list(valor.values())[0]} - Valor: ")
            valores = list(valor.values()) 
            if valores[3]=="encontrado":               
                if "Amex US$" in valores[0]:
                    total_amex += valores[2] 
                if "Dinners US$" in valores[0]:
                    total_dinners += valores[2]
                if "Master Card US$" in valores[0]:
                    total_mc += valores[2]
                if "Visa US$" in valores[0]:
                    total_visa += valores[2]
        diccionario_temp[0]["Valor US"]=total_amex
        diccionario_temp[1]["Valor US"]=total_dinners
        diccionario_temp[2]["Valor US"]=total_mc
        diccionario_temp[3]["Valor US"]=total_visa
        diccionario_temp[4]["Valor US"]=total_visa+total_amex+total_dinners+total_mc
        print(diccionario_temp)
        
        for clave, valor in diccionario_2.items():
            print(f"Clave: {clave} -> Valor: {valor} -> Dato Valor 1:{list(valor.values())[0]} - Valor: ")
            valores = list(valor.values())
            print(f"Clave:{clave} Valores:{valor}") 
            if valores[3]=="encontrado":               
                if "AX" in valores[0]:
                    total_amex += valores[2] 
                if "DI" in valores[0]:
                    total_dinners += valores[2]
                if "MC" in valores[0]:
                    total_mc += valores[2]
                if "VI" in valores[0]:
                    total_visa += valores[2]
        diccionario_temp[0]["Valor TBK"]=total_amex
        diccionario_temp[1]["Valor TBK"]=total_dinners
        diccionario_temp[2]["Valor TBK"]=total_mc
        diccionario_temp[3]["Valor TBK"]=total_visa
        diccionario_temp[4]["Valor TBK"]=total_visa+total_amex+total_dinners+total_mc
        print(diccionario_temp)
        diccionario_temp[0]["Diferencia"] = diccionario_temp[0]["Valor TBK"] - diccionario_temp[0]["Valor US"]
        diccionario_temp[1]["Diferencia"] = diccionario_temp[1]["Valor TBK"] - diccionario_temp[1]["Valor US"]
        diccionario_temp[2]["Diferencia"] = diccionario_temp[2]["Valor TBK"] - diccionario_temp[2]["Valor US"]
        diccionario_temp[3]["Diferencia"] = diccionario_temp[3]["Valor TBK"] - diccionario_temp[3]["Valor US"]
        diccionario_temp[4]["Diferencia"] = diccionario_temp[4]["Valor TBK"] - diccionario_temp[4]["Valor US"]
        
        contexto["data"]=diccionario_temp        
    else:
            print("es invalido")

    return render(request,"procesamiento.html",contexto)


data = {}
def descargar_excel_ant(request):
    # Datos de ejemplo
    data = {
        'Documento': documento_inf,
        'Monto': monto_buscar_inf,
        'Fecha': fecha_inf,
        'TC':tc_inf
    }
    df = pd.DataFrame(data)


    # Crear archivo Excel en memoria
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    # Crear respuesta HTTP para descarga
    response = HttpResponse(
        buffer,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=mi_archivo.xlsx'
    return response

def descargar_excel(request):
    # Datos de ejemplo
    data = {
        'Documento': documento_inf,
        'Monto': monto_buscar_inf,
        'Fecha': fecha_inf,
        'TC': tc_inf
    }
    df = pd.DataFrame(data)

    # Crear archivo Excel en memoria
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl', startrow=1)  # empezar en fila 2
    buffer.seek(0)

    # Abrir el workbook con openpyxl para agregar el título
    wb = load_workbook(buffer)
    ws = wb.active

    # Agregar título en la primera fila
    ws['A1'] = "Reporte de Transacciones No Encontradas"  # texto del título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=df.shape[1])  # fusionar celdas del título

    # Guardar nuevamente en el buffer
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # Crear respuesta HTTP para descarga
    response = HttpResponse(
        buffer,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=mi_archivo.xlsx'
    return response
   
def generar_qr2(req):
    # Datos a codificar
    data = "https://example.com"
    
    # Crear código QR
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill='black', back_color='white')

    # Convertir a bytes
    buf = BytesIO()
    img.save(buf, format='PNG')
    img_bytes = buf.getvalue()

    # Guardar
    qr_image = ContentFile(buf.getvalue(), 'qr_code.png')
    # Responder con imagen
    return HttpResponse(img_bytes, content_type='image/png')
     
def login(request):
    contexto = {}
    if request.POST:
        nombre = request.POST.get("email")
        password = request.POST.get("pass")
        us = authenticate(request,username=nombre,password=password)
        if us is not None and us.is_active:
            login_aut(request,us)
            usuario=User.objects.get(email=request.POST.get("email"))
            print("usuario")
            print(usuario.id)
            cli=Cliente.objects.get(email=nombre)
            request.session["email"]=nombre
            return render(request,"index.html",contexto)
        else:
            contexto = {"mensaje":"usuario y contraseña incorrecto"}
            return render(request,"login.html",contexto)        
    return render(request,"login.html",contexto)

def enviar_codigo_qr(request,clave,correo):
    
    # Datos que quieres convertir en un código QR
    data = clave  # Reemplázalo con la URL o información que desees

    # Generar el código QR
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    # Crear la imagen del código QR
    img = qr.make_image(fill_color="black", back_color="white")

    # Guardar la imagen en un buffer
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)

    # Crear el mensaje de correo electrónico
    email = EmailMessage(
        subject=correo,
        body='Adjunto encontrarás tu código QR.de tu reserva',
        from_email='fm_campos@yahoo.com',  # Reemplaza con tu correo
        to=[correo],   # Reemplaza con el correo del destinatario
    )

    # Adjuntar la imagen del código QR
    email.attach('codigo_qr.png', buffer.getvalue(), 'image/png')

    # Enviar el correo
    email.send()

    comentarios=Comentario.objects.all()
    contexto={'comentarios':comentarios}
    contexto["mensaje"]="OK"
    return 1

def cerrar_sesion(request):
    contex = {}
    logout(request)
    return render(request,"index.html",contex)
    

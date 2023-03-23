# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 14:35:54 2023

@author: Asus
"""

import pandas as pd
import glob
import numpy as np
import xlsxwriter
from datetime import timedelta

inputs = pd.read_excel('Datos_LBC.xlsx',sheet_name='festivos')
festivos = inputs['Fecha'].tolist()
df_pruebas =  pd.read_excel(r'C:\Users\Asus\OneDrive - julia-rd.com.co\Direccion MEM\Operacion_MEM\06. RESULTADOS_PRUEBAS\Consolidado_PRUEBAS.xlsx')
df_pruebas['desconexion'] = 1
df_pruebas = df_pruebas[['FronteraID','FechaOperacion','desconexion']]


ls_df = []
ls_nombres = []
for f in glob.glob("Matrices/*.xlsx"):   ###LECTURA MATRICES DE CONSUMO ACTUALES
    ls_df.append(pd.read_excel(f, sheet_name='Datos',skiprows=6,header=1))
    ls_nombres.append(f[25:33])

def fun(num):
    if num in festivos:
        return 'Festivo'
def Depuracion_datos (df,nomb):
    df['Fecha Observación'] = pd.to_datetime(df['Fecha Observación'],format='%m/%d/%Y',errors='coerce')
    df_pruebas2 = df_pruebas[df_pruebas['FronteraID'] == nomb]
    df = df.merge(df_pruebas2,left_on='Fecha Observación', right_on='FechaOperacion', how='left')
    df.drop(columns=['FechaOperacion','FronteraID'], inplace=True)
    df = df.groupby('Fecha Observación').sum()  ###SUMAR CONSUMOS DE VARIOS PREDIOS
    df = df.reset_index()
    df['ds']=df['Fecha Observación'].dt.day_name() ##COLOCAR TIPO DE DIA LUNES-DOMINGO
    df['ds2'] = df['Fecha Observación'].apply(fun) ##AGREGAR FESTIVOS
    df['Dia_semana'] = 'validar'
    for i in range(len(df)):  ###AGREGAR FESTIVOS NO DOMINGO (SE DEBE HACER UNA LISTA DE FESTIVOS QUE NO SEAN DOMINGOS)
        if df['ds'].iloc[i] == 'Sunday':
            df.at[i, 'Dia_semana'] = 'Sunday'
        elif df['ds2'].iloc[i]  == 'Festivo':
            df.at[i, 'Dia_semana'] = 'Festivo'
        else:
            df.at[i, 'Dia_semana'] = df['ds'].iloc[i] 
    dia_60 = df['Fecha Observación'].max() - timedelta(days=60) ###ACOTACION 60 DIAS
    df2 = df[(df['Fecha Observación'] > dia_60)]
    df3= df[(df['Fecha Observación'] <= dia_60)]
    if len(df2[df2['Dia_semana']=='Festivo']) < 5: ##AMPLIAR MUESTRA FESTIVOS 105 dias
        add = df3[df3['Dia_semana']=='Festivo']
        df4 = pd.concat([df2, add], axis=0)
        if len(df4[df4['Dia_semana']=='Festivo']) < 5: ### PONER DOMINGO MAS BAJOS
            add2 = df4[(df4['Dia_semana']=='Sunday')&(df4['desconexion'] != 1)]  ### agregar domingos que no sean desconexiones
            add2 = add2.sort_values(['Demanda DDV'])
            falt = 5 - len(df4[df4['Dia_semana']=='Festivo'])
            add2 = add2[0:falt]
            add2[ 'Dia_semana'] = 'Festivo'
            df = pd.concat([df4, add2], axis=0)
    else:
        df = df2    
    df = df.reset_index()
    df.drop(columns="index", inplace=True)
    return df

def Tipo_dia (num): ## CREACION COLUMNA DE DIAS TIPO NUEVA RESOLUCION
    if num=='Sunday':
        return 'Domingo'
    elif num=='Festivo':
        return 'Festivo'
    elif num=='Saturday':
        return 'Sabado'
    else:
        return 'Laboral'
def conteo_dias (df): ###CONTEO DIAS POR DIA TIPO
    dias_sem = df['Tipo_dia'].unique().tolist()
    df['conteo'] = 0
    for fildias in dias_sem:
        cont=df['Demanda DDV'][df['Tipo_dia']==fildias].count()
        df.loc[df['Tipo_dia'] == fildias, 'conteo'] = cont  ###CONTR DIAS POR DIA TIPO
    return df

def reemplazar_ceros (df): ###REEMPLAZAR CEROS CON PROMEDIO MOVIL (si el codgio aun no da igual sacar desconexiones del promedio movil)
    dias_sem = df['Tipo_dia'].unique().tolist()
    df_dias =[]
    for fildias in dias_sem:
        dff=df[df['Tipo_dia']==fildias].sort_values(by=['Fecha Observación']).reset_index()
        val = dff['Demanda DDV'].tolist()
        dff2 = df[df['Tipo_dia']==fildias].sort_values(by=['Fecha Observación']).reset_index()
        if 0 in val:
            df_ceros = dff.index[(dff['Demanda DDV']==0) & (df['desconexion']!=1)]
            delv = []
            for z in range(len(df_ceros)):
                a=df_ceros[z]
                try:
                    if a-1 >= 0:
                        if dff['Demanda DDV'].iloc[a-1] == 0:
                            zm1=0
                        else:
                            zm1 = float(dff['Demanda DDV'].iloc[a-1])
                    else:
                        zm1=0
                except:
                    zm1 = 0
                try:
                    if a-2 >= 0:
                        if dff['Demanda DDV'].iloc[a-2] == 0:
                            zm2=0
                        else:
                            zm2 = float(dff['Demanda DDV'].iloc[a-2])
                    else:
                        zm2=0
                except:
                    zm2=0
                try:
                    if dff['Demanda DDV'].iloc[a+1] == 0:
                        z1=0
                    else:
                        z1 = float(dff['Demanda DDV'].iloc[a+1])
                except:
                    z1=0
                try:
                    if dff['Demanda DDV'].iloc[a+2] == 0:
                        z2=0
                    else:
                        z2 = float(dff['Demanda DDV'].iloc[a+2])
                except:
                    z2=0
                lv = [zm1,zm2,z1,z2]
                lv = [i for i in lv if i != 0]
                if len(lv) == 0:
                    delv.append(a)
                else:
                    prom = sum(lv)/len(lv)
                    dff.loc[a,'Demanda DDV'] = prom
                    dff2.loc[a,'Demanda DDV'] = prom
            dff2=dff2.drop(dff2.index[delv])
        df_dias.append(dff2)
    df = pd.concat(df_dias)

    return df
     
def Maximo_total (df):
    dias_sem = df['Tipo_dia'].unique().tolist()
    df_dias=[]
    for fildias in dias_sem:
        if fildias == 'Festivo':
            dfp = df[df['Tipo_dia']==fildias]
            df_dias.append(dfp)
        else:      
            dfp = df[(df['Tipo_dia']==fildias) & (df['desconexion']!=1)]
            dfp = dfp.drop(dfp['Demanda DDV'].idxmax())
            df_dias.append(dfp)
            df_dias.append(df[(df['Tipo_dia']==fildias) & (df['desconexion']==1)])
    df = pd.concat(df_dias)
    return df

def Minimo_total (df):
    dias_sem = df['Tipo_dia'].unique().tolist()
    df_dias=[]
    for fildias in dias_sem:
        if fildias == 'Festivo':
            dfp = df[df['Tipo_dia']==fildias]
            df_dias.append(dfp)
        else:      
            dfp = df[(df['Tipo_dia']==fildias) & (df['desconexion']!=1)]
            dfp = dfp.drop(dfp['Demanda DDV'].idxmin())
            df_dias.append(dfp)
            df_dias.append(df[(df['Tipo_dia']==fildias) & (df['desconexion']==1)])
    df = pd.concat(df_dias)
    return df

def eliminacion_atipicos (df): ###ELIMINACION ATIPICOS (si el codgio aun no da igual sacar desconexiones del promedio movil)
    dias_sem = df['Tipo_dia'].unique().tolist()
    df_dias =[]
    for fildias in dias_sem:
        dff=df[df['Tipo_dia']==fildias].sort_values(by=['Fecha Observación']).reset_index()
        dff2 = df[df['Tipo_dia']==fildias].sort_values(by=['Fecha Observación']).reset_index() ##########'Fecha Observación'
        dff['Atipico'] = 0

        mediana = np.percentile(dff['Demanda DDV'], 50)
        q_25 = np.percentile(dff['Demanda DDV'][dff['Demanda DDV']<mediana], 50)
        q_75 = np.percentile(dff['Demanda DDV'][dff['Demanda DDV']>mediana], 50)
    
        iqr = q_75-q_25
        dff['Atipico'] = np.where((dff['desconexion']!=1) & ((dff['Demanda DDV'] < q_25-1.5*iqr) | (dff['Demanda DDV'] > q_75+1.5*iqr)), 1,0)
        
        dff = dff.drop(['level_0'], axis=1)
        val = dff['Atipico'].tolist()
        if 1 in val:
            df_atipico = dff.index[dff['Atipico'] == 1]
            delv = []
            for z in range(len(df_atipico)):
                a=df_atipico[z]
                try:
                    if a-1 >= 0:
                        if dff['Atipico'].iloc[a-1] == 1:
                            zm1=0
                        else:
                            zm1 = float(dff['Demanda DDV'].iloc[a-1])
                    else:
                        zm1=0
                except:
                    zm1 = 0
                try:
                    if a-2 >= 0:
                        if dff['Atipico'].iloc[a-2] == 1:
                            zm2=0
                        else:
                            zm2 = float(dff['Demanda DDV'].iloc[a-2])
                    else:
                        zm2=0
                except:
                    zm2=0
                try:
                    if dff['Atipico'].iloc[a+1] == 1:
                        z1=0
                    else:
                        z1 = float(dff['Demanda DDV'].iloc[a+1])
                except:
                    z1=0
                try:
                    if dff['Atipico'].iloc[a+2] == 1:
                        z2=0
                    else:
                        z2 = float(dff['Demanda DDV'].iloc[a+2])
                except:
                    z2=0
                lv = [zm1,zm2,z1,z2]
                lv = [i for i in lv if i != 0]
                if len(lv) == 0:
                    delv.append(a)
                else:
                    prom = sum(lv)/len(lv)
                    dff.loc[a,'Demanda DDV'] = prom
                    dff.loc[a,'Atipico'] = 0
                    dff2.loc[a,'Demanda DDV'] = prom
            dff2=dff2.drop(dff2.index[delv])
        df_dias.append(dff2)
    df = pd.concat(df_dias)
    return df

def transformacion_desconexiones(df):
    dias_sem = df['Tipo_dia'].unique().tolist()
    df_dias =[]
    df = df.drop(['level_0'], axis=1)
    for fildias in dias_sem:
        dff=df[df['Tipo_dia']==fildias].sort_values(by=['Fecha Observación']).reset_index()
        dff2 = df[df['Tipo_dia']==fildias].sort_values(by=['Fecha Observación']).reset_index() ##########'Fecha Observación'
        val = dff['desconexion'].tolist()
        
        if 1 in val:
            df_atipico = dff.index[dff['desconexion'] == 1]
            delv = []
            for z in range(len(df_atipico)):
                a=df_atipico[z]
                try:
                    if a-1 >= 0:
                        if dff['desconexion'].iloc[a-1] == 1:
                            zm1=0
                        else:
                            zm1 = float(dff['Demanda DDV'].iloc[a-1])
                    else:
                        zm1=0
                except:
                    zm1 = 0
                try:
                    if a-2 >= 0:
                        if dff['desconexion'].iloc[a-2] == 1:
                            zm2=0
                        else:
                            zm2 = float(dff['Demanda DDV'].iloc[a-2])
                    else:
                        zm2=0
                except:
                    zm2=0
                try:
                    if a-3 >= 0:
                        if dff['desconexion'].iloc[a-3] == 1:
                            zm3=0
                        else:
                            zm3 = float(dff['Demanda DDV'].iloc[a-3])
                    else:
                        zm3=0
                except:
                    zm3=0
                try:
                    if a-4 >= 0:
                        if dff['desconexion'].iloc[a-4] == 1:
                            zm4=0
                        else:
                            zm4 = float(dff['Demanda DDV'].iloc[a-4])
                    else:
                        zm4=0
                except:
                    zm4=0
                lv = [zm1,zm2,zm3,zm4]
                lv = [i for i in lv if i != 0]
                if len(lv) == 0:
                    try:
                        if a+1 >= 0:
                            if dff['desconexion'].iloc[a+1] == 1:
                                zm1=0
                            else:
                                zm1 = float(dff['Demanda DDV'].iloc[a+1])
                        else:
                            zm1=0
                    except:
                        zm1 = 0
                    try:
                        if a+2 >= 0:
                            if dff['desconexion'].iloc[a+2] == 1:
                                zm2=0
                            else:
                                zm2 = float(dff['Demanda DDV'].iloc[a+2])
                        else:
                            zm2=0
                    except:
                        zm2=0
                    try:
                        if a-3 >= 0:
                            if dff['desconexion'].iloc[a+3] == 1:
                                zm3=0
                            else:
                                zm3 = float(dff['Demanda DDV'].iloc[a+3])
                        else:
                            zm3=0
                    except:
                        zm3=0
                    try:
                        if a-4 >= 0:
                            if dff['desconexion'].iloc[a+4] == 1:
                                zm4=0
                            else:
                                zm4 = float(dff['Demanda DDV'].iloc[a+4])
                        else:
                            zm4=0
                    except:
                        zm4=0
                    lv = [zm1,zm2,zm3,zm4]
                    lv = [i for i in lv if i != 0]
                    if len(lv) == 0:
                        delv.append(a)
                    else:
                        
                        prom = sum(lv)/len(lv)
                        dff.loc[a,'Demanda DDV'] = prom
                        dff.loc[a,'desconexion'] = 0
                        dff2.loc[a,'Demanda DDV'] = prom
                else:
                    
                    prom = sum(lv)/len(lv)
                    dff.loc[a,'Demanda DDV'] = prom
                    dff.loc[a,'desconexion'] = 0
                    dff2.loc[a,'Demanda DDV'] = prom
            dff2=dff2.drop(dff2.index[delv])
        df_dias.append(dff2)
    df = pd.concat(df_dias)
    return df

def lbc_final (lbc):
    lbc['LBC final'] = 0
    for i in range(len(lbc)):
        if lbc['Error RRMSE %'].iloc[i]<=5:
            lbc.at[lbc.index[i], 'LBC final'] = lbc['LBC estimada'].iloc[i]
        elif lbc['Error RRMSE %'].iloc[i]>5 and lbc['Error RRMSE %'].iloc[i]<=20:
            lbc.at[lbc.index[i], 'LBC final'] = lbc['LBC estimada'].iloc[i]*(1-lbc['Error RRMSE %'].iloc[i]/100)
        else:
            lbc.at[lbc.index[i], 'LBC final'] = 0
    return lbc

def main_error(ls_df,ls_nombres):
    lbcs = []
    for i in range(len(ls_df)):  ## LECTURA DE CADA MATRIZ
        df = ls_df[i] ##DEFINICION DE UNA MATRIZ
        nomb = ls_nombres[i]
        df = Depuracion_datos(df,nomb)  ###AGREGAR TIPO DIA NUEVA RESOLUCION Y LIMITAR MUESTRA A 60 DIA
        df['Tipo_dia'] = df['Dia_semana'].apply(Tipo_dia) ##TIPO DIAS
        df2 = conteo_dias(df)
        df = conteo_dias(df) ###CONTAR CUANTOS DIAS HAY POR DIA TIPO
        df = reemplazar_ceros(df)  ###REEMPLAZAR CEROS CON PROMEDIO MOVIL
        df = Maximo_total(df) ###SACAR MAXIMO EXCEPTO PARA FESTIVO
        df = Minimo_total(df) ### SACAR MINIMO EXCEPTO PARA FESTIVO
        df = eliminacion_atipicos(df)  ###ELIMINACION ATIPICOS
        df = transformacion_desconexiones(df) ###ELIMINACION ATIPICOS
        df.drop(columns=['ds', 'ds2','Unnamed: 0','Demanda Diaria por Frontera'], inplace=True)
        lbc = pd.pivot_table(data=df,index=['Tipo_dia'],values=['Demanda DDV'],aggfunc='mean')
        conteo = pd.pivot_table(data=df2, index=['Tipo_dia'], values='conteo',aggfunc='count')
        sum_dia = pd.pivot_table(data=df2, index=['Tipo_dia'], values='Demanda DDV',aggfunc='sum')
        sum_dia = sum_dia.rename(columns={'Demanda DDV': 'Sum Ci',})
        lbc = pd.merge(lbc, conteo, left_index=True, right_index=True)
        lbc = pd.merge(lbc, sum_dia, left_index=True, right_index=True)
        lbc = lbc.rename(columns={'conteo': 'Conteo'})
        lbc['1/n'] =lbc['Conteo'].apply(lambda x: 1/x)
        lbc = lbc.rename(columns={'Demanda DDV': 'LBC estimada'})
        df2 = df2.merge(lbc[['LBC estimada']],left_on='Tipo_dia', right_index=True, how='left')
        df2['(LBC - Ci)^2'] = (df2['LBC estimada']-df2['Demanda DDV'])**2
        sum_lbc_ci = pd.pivot_table(data=df2, index=['Tipo_dia'], values='(LBC - Ci)^2',aggfunc='sum')
        sum_lbc_ci = sum_lbc_ci.rename(columns={'(LBC - Ci)^2': 'sum (LBC - Ci)^2'})
        lbc = pd.merge(lbc, sum_lbc_ci, left_index=True, right_index=True)
        lbc['Error RRMSE %'] = (((lbc['1/n']*lbc['sum (LBC - Ci)^2'])**0.5)/(lbc['1/n']*lbc['Sum Ci']))*100
        lbc['Error RRMSE %'] = lbc['Error RRMSE %'].apply(lambda x: round(x,2))
        lbc = lbc_final(lbc)
        lbcs.append(lbc)
    return lbcs

def crear_excel_salida(lbcs,ls_nombres):
    workbook = xlsxwriter.Workbook('LBC_Y_ERROR.xlsx')
    worksheet = workbook.add_worksheet('LBC')
    
    merge_format1 = workbook.add_format({'border': 1,}) 
    merge_format2 = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#99CCFF'})
    merge_format3 = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#FFFFFF'})
    merge_format4 = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#FF0000'})
    merge_format5 = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#dfa801'})  
    merge_format6 = workbook.add_format({'bold': 1,'border': 1,'align': 'center','valign': 'vcenter','fg_color': '#F79441'})
    for i in range(len(lbcs)):
        df_lbc = lbcs[i]
        nombre = ls_nombres[i]
        worksheet.write(i+1,0,nombre,merge_format1)
        dias_sem = df_lbc.index.tolist()
        for k in range(len(dias_sem)):
            worksheet.write(0,1,'LBC GORRO LABORAL',merge_format2)
            worksheet.write(0,2,'LBC GORRO DOMINGO',merge_format2)
            worksheet.write(0,3,'LBC GORRO SABADO',merge_format2)
            worksheet.write(0,4,'LBC GORRO FESTIVO',merge_format2)
            error = df_lbc['Error RRMSE %'][df_lbc.index==dias_sem[k]]
            error = float(error)
            lbcw = df_lbc['LBC final'][df_lbc.index==dias_sem[k]]
            lbce = df_lbc['LBC estimada'][df_lbc.index==dias_sem[k]]
            
            if dias_sem[k] == 'Laboral':
                worksheet.write(0,1+4,'ERROR LABORAL',merge_format2)
                worksheet.write(0,5+4,'LBC LABORAL',merge_format6)
                worksheet.write(i+1,1,lbce,merge_format1)
                if error > 20:
                    worksheet.write(i+1,1+4,error,merge_format4)
                    worksheet.write(i+1,5+4,lbcw,merge_format4)
                elif error > 5 and error <= 20:
                    worksheet.write(i+1,1+4,error,merge_format5)
                    worksheet.write(i+1,5+4,lbcw,merge_format5)
                else:
                    worksheet.write(i+1,1+4,error,merge_format3)
                    worksheet.write(i+1,5+4,lbcw,merge_format3)
                
            elif dias_sem[k] =='Domingo':
                worksheet.write(0,2+4,'ERROR DOMINGO',merge_format2)
                worksheet.write(0,6+4,'LBC DOMINGO',merge_format6)
                worksheet.write(i+1,2,lbce,merge_format1)
                if error > 20:
                    worksheet.write(i+1,2+4,error,merge_format4)
                    worksheet.write(i+1,6+4,lbcw,merge_format4)
                elif error > 5 and error <= 20:
                    worksheet.write(i+1,2+4,error,merge_format5)
                    worksheet.write(i+1,6+4,lbcw,merge_format5)
                else:
                    worksheet.write(i+1,2+4,error,merge_format3)
                    worksheet.write(i+1,6+4,lbcw,merge_format3)
            elif dias_sem[k] =='Sabado':
                worksheet.write(0,3+4,'ERROR SABADO',merge_format2)
                worksheet.write(0,7+4,'LBC SABADO',merge_format6)
                worksheet.write(i+1,3,lbce,merge_format1)
                if error > 20:
                    worksheet.write(i+1,3+4,error,merge_format4)
                    worksheet.write(i+1,7+4,lbcw,merge_format4)
                elif error > 5 and error <= 20:
                    worksheet.write(i+1,3+4,error,merge_format5)
                    worksheet.write(i+1,7+4,lbcw,merge_format5)
                else:
                    worksheet.write(i+1,3+4,error,merge_format3)
                    worksheet.write(i+1,7+4,lbcw,merge_format3)
            elif dias_sem[k] =='Festivo':
                worksheet.write(0,4+4,'ERROR FESTIVO',merge_format2)
                worksheet.write(0,8+4,'LBC FESTIVO',merge_format6)
                worksheet.write(i+1,4,lbce,merge_format1)
                if error > 20:
                    worksheet.write(i+1,4+4,error,merge_format4)
                    worksheet.write(i+1,8+4,lbcw,merge_format4)
                elif error > 5 and error <= 20:
                    worksheet.write(i+1,4+4,error,merge_format5)
                    worksheet.write(i+1,8+4,lbcw,merge_format5)
                else:
                    worksheet.write(i+1,4+4,error,merge_format3)
                    worksheet.write(i+1,8+4,lbcw,merge_format3)

    worksheet.write(0,0,'FRONTERA',merge_format2) 
    workbook.close()

lbcs = main_error(ls_df,ls_nombres)  
crear_excel_salida(lbcs,ls_nombres)
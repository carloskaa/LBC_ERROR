# -*- coding: utf-8 -*-
"""
Created on Wed Mar 22 19:26:40 2023

@author: Asus
"""

import zipfile
import pandas as pd
import streamlit as st
import numpy as np
from XM_LBC_VERSION_FINAL import crear_excel_salida
from XM_LBC_VERSION_FINAL import main_error
import io
st.title("APLICACION MARGEN DE ERROR")
uploaded_file = st.file_uploader("Seleccione un archivo .zip", type="zip")
val = 0
val2 = 0
if uploaded_file is not None:
    ls_df = []
    ls_nombres = []
    with zipfile.ZipFile(uploaded_file) as zip_file:
        for file in zip_file.namelist():
            # st.write(f"Archivo: {file}")
            if file.endswith(".xlsx"):
                st.write(f"Leyendo archivo de Excel: {file}")
                ls_df.append(pd.read_excel(zip_file.open(file), dtype={'Demanda DDV': np.float64, 'Demanda Diaria por Frontera':np.float64},sheet_name='Datos',skiprows=6,header=1))
                ls_nombres.append(file[25:33])
            else:
                file_contents = zip_file.read(file)
                # st.write(file_contents)
        val =1

def generar_boton_descarga():
    output = io.BytesIO()
    with open('LBC_Y_ERROR.xlsx', 'rb') as archivo:
        contenido = archivo.read()
    output.write(contenido)
    output.seek(0)
    st.download_button(label='Descargar archivo', data=output, file_name='LBC_Y_ERROR.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if val == 1:
    st.success('Calculo realizado correctamente')
    val2 = 1
    
if val2 == 1:
    if st.checkbox("Ver resultados"):
        lbcs = main_error(ls_df,ls_nombres) 
        for i in lbcs:
            st.dataframe(i)
        crear_excel_salida(lbcs,ls_nombres)
        generar_boton_descarga()

















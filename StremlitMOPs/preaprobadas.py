import streamlit as st
import openpyxl

def mopPreaprobadas():

    workbook = openpyxl.load_workbook('MopsPreaprobadas.xlsx')
    sheets = workbook.sheetnames
    sheet = workbook.active

    objetivo = st.text_input('Objetivo:')
    sheet['C7'] = objetivo

    nomMantenimiento = st.text_input('Nombre de mantenimiento:')
    sheet['C8'] = nomMantenimiento

    descripcion = st.text_input('Descripción de impacto:', value='No se Espera Afectación')
    sheet['C9'] = descripcion

    col1, col2 = st.columns(2)
    with col1:
        fechaIniciomtto = st.date_input('Fecha Inicio MTTO')
        sheet['F13'] = fechaIniciomtto.strftime('%d-%b-%Y')
        fechaFinmtto = st.date_input('Fecha Fin MTTO')
        sheet['F14'] = fechaIniciomtto.strftime('%d-%b-%Y')
        fechaInicioAfectacion = st.date_input('Fecha Inicio Afectación')
        sheet['F15'] = fechaInicioAfectacion.strftime('%d-%b-%Y')
        fechaFinAfectacion = st.date_input('Fecha Fin Afectación')
        sheet['F16'] = fechaFinAfectacion.strftime('%d-%b-%Y')
        fechaInicioRollback = st.date_input('Fecha Inicio Rollback')
        sheet['F17'] = fechaInicioRollback.strftime('%d-%b-%Y')
        fechaFinRollback = st.date_input('Fecha Fin Rollback')
        sheet['F18'] = fechaFinRollback.strftime('%d-%b-%Y')

    with col2:
        horaIniciomtto = st.text_input('Hora Inicio MTTO', value='00:00')
        sheet['G13'] = horaIniciomtto
        horaFinmtto = st.text_input('Hora Fin MTTO', value='05:00')
        sheet['G14'] = horaFinmtto
        horaInicioAfectacion = st.text_input('Hora Inicio Afectacion', value='00:00')
        sheet['G15'] = horaInicioAfectacion
        horaFinAfectacion = st.text_input('Hora Fin Afectacion',value='00:00')
        sheet['G16'] = horaFinAfectacion
        horaInicioRollback = st.text_input('Hora Inicio Rollback', value='03:30')
        sheet['G17'] = horaInicioRollback
        horaFinRollback = st.text_input('Hora Fin Rollback',value='04:30')
        sheet['G18'] = horaFinRollback

        sheet['C13'] = str(fechaIniciomtto.strftime('%d-%m-%Y')) +" "+ str(horaIniciomtto)
        sheet['C14'] = str(fechaFinmtto.strftime('%d-%m-%Y')) +" "+ str(horaFinmtto)
        sheet['C15'] = str(fechaInicioAfectacion.strftime('%d-%m-%Y')) +" "+ str(horaInicioAfectacion)
        sheet['C16'] = str(fechaFinAfectacion.strftime('%d-%m-%Y')) +" "+ str(horaFinAfectacion)
        sheet['C17'] = str(fechaInicioRollback.strftime('%d-%m-%Y')) +" "+ str(horaInicioRollback)
        sheet['C18'] = str(fechaFinRollback.strftime('%d-%m-%Y')) +" "+ str(horaFinRollback)

    ppmProyecto = st.text_input('PPM (Proyecto)',value='JIRA-')
    sheet['C19'] = ppmProyecto
    col12,col22 = st.columns(2)
    with col12:
        resOtecel = st.text_input('Responsable Otecel', value='N2 O&M')
        sheet['C20'] = resOtecel
        resProveedor = st.text_input('Responsable Proveedor', value='N1 PAP')
        sheet['C23'] = resProveedor
        

    with col22:
        tlfOtecel = st.text_input('Telefono Otecel', value='0999729993')
        if tlfOtecel.isdigit() == True:
            sheet['C21'] = tlfOtecel
        else:
            st.warning('Ingrese un numero valido')

        tlfProveedor = st.text_input('Telefono Proveedor', value='0999993070')
        if tlfProveedor.isdigit() == True:
            sheet['C24'] = tlfProveedor
        else:
            st.warning('Ingrese un numero valido')
        
        emailProveedor = st.text_input('Email Proveedor', value='n1pap.ec@telefonica.com')
        sheet['C25'] = emailProveedor

    if st.button('Guardar'):
        workbook.save('MOP_Preaprobadas.xlsx')

    st.subheader('Descargar Archivo')
    if st.checkbox('Descargar:'):
        with open('MOP_Preaprobadas.xlsx','rb') as fp:
            btn = st.download_button(label='Descargar Archivo', data=fp, file_name='MOP_Preaprobadas.xlsx')

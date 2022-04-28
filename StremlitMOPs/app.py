# Bibliotecas
import streamlit as st
import openpyxl
from preaprobadas import mopPreaprobadas
from mantenimientos import mopMantenimientos

def main():

    #st.image('proconty.png', use_column_width=True)
    st.title('PROCONTY - Pasos a Producción')
    
    st.sidebar.image('proconty.png', width=300)
    
    menu = ['MOP Gestión', 'MOP Preaprobadas']
    choice = st.sidebar.selectbox('Menu', menu)

    if choice == 'MOP Gestión':
        st.subheader('MOP Gestión')
        mopMantenimientos()

    elif choice == 'MOP Preaprobadas':
        st.subheader('MOP Preaprobadas')
        mopPreaprobadas()


    st.sidebar.markdown('[PROCONTY](https://www.facebook.com/PROCONTY/)')
    st.sidebar.write('All rights reserved. Developed by David Minango')
    

if __name__ == '__main__':
    main()
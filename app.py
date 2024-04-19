import streamlit as st
import src.datos as datos


st.title('Movie App')

st.header('El dataset utilizado es el siguiente:')
st.write(datos.leer_datos())

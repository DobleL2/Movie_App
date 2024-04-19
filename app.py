import streamlit as st
import src.datos as datos
import src.graficos as graficos
import src.analitica as analitica


st.title('Movie App')

st.header('El dataset utilizado es el siguiente:')

st.write(datos.leer_datos())

graficos.resumen_general()

year = st.selectbox('Año: ',options= datos.retornar_years() )

st.subheader(f'Top 10 películas {year}')

col1,col2 = st.columns(2)

col1.write(analitica.top_10_year(year))

graficos.graficos_dona(year,col2)

graficos.graficos_histo(year)
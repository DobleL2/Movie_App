import streamlit as st

def graficos_dona(year,col):

    col.image(f'images/donuts_platform_year_{year}.png',caption=f'Porcentajes de peliculas en cada plataforma en el año {year}')

def graficos_histo(year):

    st.image(f'images/histogram_age_year_{year}.png',caption=f'Histograma de edades permitidas el año {year}')
    
def resumen_general():
    st.image(f'images/visualization.png',caption='Resultados generales resumidos por años y edades permitidas')
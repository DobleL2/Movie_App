import streamlit as st
import src.datos as datos
import src.graficos as graficos
import src.analitica as analitica
import src.reporte_pdf as reporte_pdf
import src.reporte_excel as reporte_excel


st.title('Movie App')
st.divider()
st.header('El dataset utilizado es el siguiente:')

st.write(datos.leer_datos())

graficos.resumen_general()
st.divider()
year = st.selectbox('Año: ',options= datos.retornar_years() )

st.subheader(f'Top 10 películas {year}')

col1,col2 = st.columns(2)

col1.write(analitica.top_10_year(year))

graficos.graficos_dona(year,col2)

graficos.graficos_histo(year)

pdf = reporte_pdf.crear_reporte(year)
st.divider()
col_1,col_2 = st.columns(2)

html = reporte_pdf.create_download_link(pdf.output(dest="S").encode("latin-1"), f"Reporte_{year}")
col_1.markdown(html, unsafe_allow_html=True)

wb = reporte_excel.crear_reporte_excel(year)
excel_file = reporte_excel.descargar_excel(wb)
col_2.download_button(
    label="Descargar Reporte en Excel",
    data=excel_file,
    file_name=f"Reporte_{year}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
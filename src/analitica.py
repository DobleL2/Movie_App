import src.datos as datos

def edades_permitidas(year):
    base = datos.leer_datos()
    base_filtrada_year = base[base['Year']==year]

    cantidad_year = base_filtrada_year.groupby(by='Age').count()[['ID']].reset_index()
    cantidad_year.columns = ['Age','Amount']
    return cantidad_year

def top_10_year(year):
    base = datos.leer_datos()
    base['Ranking'] = base['Rotten Tomatoes'].str.split('/',expand=True)[0].fillna(0).astype(int)
    
    base_filtrada_year = base[base['Year']==year]
    
    return base_filtrada_year.sort_values(by='Ranking',ascending=False)[:10][['Title','Age']].reset_index(drop=True)

def plataforma(year):
    base = datos.leer_datos()
    base_filtrada_year = base[base['Year']==year]
    
    return base_filtrada_year[['Netflix','Hulu','Prime Video','Disney+']].sum()
import pandas as pd

def leer_datos():
    base = pd.read_csv('Datasets/MoviesOnStreamingPlatforms.csv')
    return base

def retornar_years():
    base = leer_datos()
    base = base.sort_values(by='Year',ascending=False)
    base['Year'] = base['Year'].astype(int) 
    years = list(base['Year'].unique())
    return years
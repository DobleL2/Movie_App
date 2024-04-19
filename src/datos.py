import pandas as pd

def leer_datos():
    base = pd.read_csv('Datasets/MoviesOnStreamingPlatforms.csv')
    return base
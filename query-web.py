import streamlit as st
import pandas as pd
import pickle
import os
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import nltk

nltk.download('stopwords')

SCOPES = ['https://www.googleapis.com/auth/webmasters.readonly']
CREDENTIALS_FILE = 'client_secret.json'
URI_REDIRECCIONAMIENTO = [
    'http://localhost:8501/',
    'https://pntc-query.streamlit.app'
     ]
VIEW_ID = 'https://pintuco.com.co/'


# Función para guardar las credenciales en un archivo pickle
def guardar_credenciales(credenciales):
    with open('credenciales.pickle', 'wb') as f:
        pickle.dump(credenciales, f)


# Función para cargar las credenciales desde un archivo pickle
def cargar_credenciales():
    try:
        with open('credenciales.pickle', 'rb') as f:
            return pickle.load(f)
    except FileNotFoundError:
        return None


# Función para autenticar o cargar las credenciales
def autenticar():
    credenciales = cargar_credenciales()

    if credenciales:
        return build('webmasters', 'v3', credentials=credenciales)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES,
                                                         redirect_uri=URI_REDIRECCIONAMIENTO)
        credenciales = flow.run_local_server(port=55875)
        guardar_credenciales(credenciales)
        return build('webmasters', 'v3', credentials=credenciales)


def obtener_datos_rendimiento(servicio, start_date, end_date):
    return servicio.searchanalytics().query(
        siteUrl=VIEW_ID,
        body={
            'startDate': start_date,
            'endDate': end_date,
            'dimensions': ['page'],
            'dimensionFilterGroups': [{
                'filters': [{
                    'dimension': 'page',
                    'operator': 'contains',
                    'expression': '/blog/'
                }]
            }],
            'metrics': ['clicks', 'impressions', 'ctr', 'position'],
            'rowLimit': 3000
        }
    ).execute()

def obtener_palabras_clave(url):
    # Eliminar stopwords
    stop_words = set(stopwords.words('spanish'))

    # Obtener las palabras clave de la URL
    palabras_clave = [word for word in url.split('/') if word.lower() not in stop_words]

    return palabras_clave


# Exportar primero la consulta
def exportar_consulta_base(datos_rendimiento, start_date_str, end_date_str, urls_especificas, keywords_especificas):
    # Obtener los datos de la consulta base
    filas = datos_rendimiento['rows']
    datos_base = []
    urls = []
    print('query OK')

    for fila in filas:
        url = fila['keys'][0]
        clicks = fila['clicks']
        impressions = fila['impressions']
        ctr = fila['ctr']
        position = fila['position']
        urls.append(url)
        datos_base.append({'URL': url, 'Clicks': clicks, 'Impressions': impressions, 'CTR': ctr, 'Position': position})

    # Crear DataFrame con los datos de la consulta base
    df = pd.DataFrame(datos_base)

    # Realizar la extracción de frases clave y análisis de similitud de texto
    vectorizer = TfidfVectorizer()
    matriz_tfidf = vectorizer.fit_transform(urls)
    similitud = cosine_similarity(matriz_tfidf)

    # Agregar las frases clave y sumar las métricas de los conjuntos similares
    conjuntos = {}
    for i, url_especifica in enumerate(urls_especificas):
        # Calcular la similitud con las URLs en los conjuntos generados
        tfidf_url_especifica = vectorizer.transform([url_especifica])
        similitudes = cosine_similarity(tfidf_url_especifica, matriz_tfidf)

        # Encontrar el conjunto más similar
        indice_similar = similitudes.argmax()
        url_similar = urls[indice_similar]

        conjunto_similar = None
        for conjunto, urls_conjunto in conjuntos.items():
            if url_similar in urls_conjunto:
                conjunto_similar = conjunto
                break

        if conjunto_similar:
            # Si se encontró un conjunto similar, agregar la URL al conjunto
            conjuntos[conjunto_similar].append(url_similar)
        else:
            # Si no se encontró un conjunto similar, crear uno nuevo
            palabras_clave = obtener_palabras_clave(url_similar) # Usar url_similar en lugar de url
            conjunto_nuevo = ' '.join(palabras_clave)
            conjuntos[conjunto_nuevo] = [url_similar]

    # Crear un nuevo DataFrame con los conjuntos y sus métricas
    conjuntos_datos = []
    for conjunto, urls_conjunto in conjuntos.items():
            clicks_suma = df.loc[df['URL'].isin(urls_conjunto), 'Clicks'].sum()
            impressions_suma = df.loc[df['URL'].isin(urls_conjunto), 'Impressions'].sum()
            ctr_suma = (df.loc[df['URL'].isin(urls_conjunto), 'CTR'] * df.loc[df['URL'].isin(urls_conjunto), 'Impressions']).sum()
            ctr_agregado = ctr_suma / impressions_suma
            position_promedio = (df.loc[df['URL'].isin(urls_conjunto), 'Position'] * df.loc[df['URL'].isin(urls_conjunto), 'Impressions']).sum() / impressions_suma
            conjuntos_datos.append({'Frase clave': conjunto, 'Clicks': clicks_suma, 'Impressions': impressions_suma, 'CTR': ctr_agregado, 'Position': position_promedio})

    df_conjuntos = pd.DataFrame(conjuntos_datos)

    # Formatear CTR y Position
    df_conjuntos['CTR'] = df_conjuntos['CTR'].apply(lambda x: '{:.2%}'.format(x))
    df_conjuntos['Position'] = df_conjuntos['Position'].apply(lambda x: '{:.2f}'.format(x))

    # Reemplazar los espacios por "/" en los datos de la columna "Frase clave"
    df_conjuntos['Frase clave'] = df_conjuntos['Frase clave'].str.replace(' ', '/')

    # Guardar el DataFrame de los conjuntos en un archivo Excel
    nombre_archivo_excel = f'consulta-{start_date_str}-to-{end_date_str}.xlsx'
    df_conjuntos.to_excel(nombre_archivo_excel, index=False)
    return nombre_archivo_excel


def cargar_urls_y_keywords_desde_excel(archivo_excel):
    df = pd.read_excel(archivo_excel)
    urls = df['URL'].tolist()
    keywords = df['KEYWORD'].tolist()
    return urls, keywords

def main():
    st.title("GSC Query Exporter - Pintuco")

    start_date = st.date_input("Start Date", pd.to_datetime('2023-06-01'))
    end_date = st.date_input("End Date", pd.to_datetime('2023-06-30'))
    archivo_excel = st.file_uploader("Upload Excel file with URLs", type=['xlsx'])

    if archivo_excel and st.button("Export Data"):
        urls_especificas, keywords_especificas = cargar_urls_y_keywords_desde_excel(archivo_excel)
        servicio = autenticar()

        start_date_str = start_date.strftime('%Y-%m-%d')
        end_date_str = end_date.strftime('%Y-%m-%d')

        datos_rendimiento = obtener_datos_rendimiento(servicio, start_date_str, end_date_str)
        nombre_archivo_excel = exportar_consulta_base(datos_rendimiento, start_date_str, end_date_str, urls_especificas, keywords_especificas)

        # Proporcionar un enlace de descarga para el archivo Excel
        st.markdown(f"**Download the Resulting Excel File**")
        st.markdown(get_download_link(nombre_archivo_excel), unsafe_allow_html=True)

def get_download_link(nombre_archivo_excel):
    with open(nombre_archivo_excel, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:file/xlsx;base64,{b64}" download="{nombre_archivo_excel}">Click here to download the Excel file</a>'
    return href

if __name__ == "__main__":
    main()

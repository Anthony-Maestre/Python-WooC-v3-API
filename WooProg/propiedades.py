from woocommerce import API
import pandas as pd

wcapi = API(
url= "",
consumer_key= "",
consumer_secret= "",
wp_api=True,
timeout=180,
version="wc/v3"
)

def atributos():
    r = wcapi.get("products/attributes")
    df = pd.read_json(r.text)
    df = df[['id','name']]
    return df

def categorias():
    r = wcapi.get("products/categories")
    df = pd.read_json(r.text)
    df = df[['id','name']]
    return df

def etiquetas():
    r = wcapi.get("products/tags")
    df = pd.read_json(r.text)
    df = df[['id','name']]
    return df

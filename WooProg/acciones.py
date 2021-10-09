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


def extraer():
    r = wcapi.get("products")
    df = pd.read_json(r.text)

    df2 = pd.DataFrame()
    n = 0
    while n <= len(df):
        p = wcapi.get(f"products/{df.id[n]}/variations")
        dfp = pd.read_json(p.text)
        dfp['name'] = df.name[n]
        dfp['parId'] = df.id[n]
        df2 = pd.concat([dfp, df2], sort=False)
        n += 1

    df2.sort_values(by=['name'], inplace=True)
    df2.reset_index(inplace=True)

    writer = pd.ExcelWriter('productosint.xlsx', engine= 'xlsxwriter')
    df2[['parId', 'id', 'name', 'weight', 'price']].to_excel(writer, index=False, sheet_name='Productos en tienda')

    workbook = writer.book
    worksheet = writer.sheets['Productos en tienda']

    worksheet.set_column('B:C', 18)
    writer.save()

    return


def actualizar():
    try:
        newprod = pd.read_excel('productosint.xlsx', dtype=str)
    except:
        return
    newprod.columns = ['parId', 'id', 'name', 'weight', 'price']

    datas = {}
    n = 0
    while n < len(newprod):
        datas[n] = {'regular_price':newprod.price[n]}
        n += 1 
        
    n = 0
    while n <= len(newprod):
        wcapi.put(f"products/{newprod.parId[n]}/variations/{newprod.id[n]}", datas[n]).json()
        n += 1

    return

def crear(data):
    wcapi.post("products", data).json()

    return
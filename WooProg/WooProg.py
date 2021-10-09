import PySimpleGUI as sg
import xlsxwriter
import propiedades

sg.theme("DarkGrey3")

cats = propiedades.categorias()
atribs = propiedades.atributos()

sg.popup_no_wait('Iniciando...', non_blocking = True, button_type = 5)

def woo_prods():
    layout = [[sg.Text('Qué quiere realizar?')],
              [sg.Button('Extraer'), sg.Button('Actualizar'), sg.Button('Crear')]]

    window = sg.Window('Seleccione', layout)
    event, values = window.read()
    window.close()

    try:
        if event == 'Actualizar':
            sg.popup_no_wait('Actualizando productos...', non_blocking = True)

            try:
                newprod = pd.read_excel('productosint.xlsx', dtype=str)
            except:
                sg.popup_error('Es necesario ejecutar "Extraer" antes de actualizar')
                return
            newprod.columns = ['parId', 'id', 'name', 'weight', 'price']

            datas = {}
            n = 0
            while n < len(newprod):
                datas[n] = {'regular_price':newprod.price[n]}
                n += 1
            else:
                print('R')
            
            n = 0
            while n <= len(newprod):
                wcapi.put(f"products/{newprod.parId[n]}/variations/{newprod.id[n]}", datas[n]).json()
                print('Se ha modificado el producto: ' + newprod.name[n])
                print(wcapi.put(f"products/{newprod.parId[n]}/variations/{newprod.id[n]}", datas[n]).json())
                n += 1
            else:
                sg.popup_no_wait('Se han actualizado todos los productos', non_blocking = True)

        elif event == 'Extraer':
            sg.popup_no_wait('Extrayendo productos...', non_blocking = True)

            r = wcapi.get("products", params={"per_page": 90})
            df = pd.read_json(r.text)

            df2 = pd.DataFrame()
            n = 0
            sg.popup_no_wait('Extrayendo variaciones...', non_blocking = True)
            while n <= len(df):
                p = wcapi.get(f"products/{df.id[n]}/variations")
                dfp = pd.read_json(p.text)
                dfp['name'] = df.name[n]
                dfp['parId'] = df.id[n]
                df2 = pd.concat([dfp, df2], sort=False)
                n += 1
            else:
                sg.popup_no_wait('Guardando resultados...', non_blocking = True)

            df2.sort_values(by=['name'], inplace=True)
            df2.reset_index(inplace=True)
            df2[['id', 'name', 'weight', 'price']]

            writer = pd.ExcelWriter('productosint.xlsx', engine= 'xlsxwriter')
            df2[['parId', 'id', 'name', 'weight', 'price']].to_excel(writer, index=False, sheet_name='Productos en tienda')

            workbook = writer.book
            worksheet = writer.sheets['Productos en tienda']

            worksheet.set_column('B:C', 18)
            writer.save()
            sg.popup_no_wait('Se han guardado los resultados', non_blocking = True)

        else:
            layoutc = [[sg.Text('* Nombre del producto:'), sg.Input(key='name')],
                      [sg.Text('* Tipo de producto:'), sg.Radio('Simple', 'RADIO1', key='tipo'), sg.Radio('Variable', 'RADIO1')],
                      [sg.Text('* Precio:'), sg.Input('00.00', key='price')],
                      [sg.Text('Descripción'), sg.Multiline(size=(20,5), key='descr')],
                      [sg.OK(), sg.Cancel()]]

            windowc = sg.Window('Nuevo producto', layoutc)
            e, v = windowc.read()
            windowc.close()

            chklist = [[sg.Text('Categorias: ')]]
            n = 0
            while n < len(cats): 
                chklist.append([sg.Checkbox(cats.name[n], key=int(cats.id[n]))])
                n += 1

            if v['tipo'] == True:
                catlay = [[sg.Text('Asigna categorias a este producto')],
                          chklist]

                wincats = sg.Window('Nuevo producto', catlay)
                ec, vc = wincats.read()
                wincats.close()

                pcats = [{'id': i} for i in vc if vc[i] == True]

                data = {
                "name": v['name'],
                "type": "simple",
                "regular_price": v['price'],
                "description": v['descr'],
                "categories": pcats
                }

                wcapi.post("products", data).json()
                sg.popup_no_wait('Producto creado con exito', non_blocking = True)

            else:
                chklist2 = [[sg.Text('Atributos: ')]]
                n = 0
                while n < len(atribs): 
                    chklist2.append([sg.Checkbox(atribs.name[n], key=int(atribs.id[n]+1000))])
                    n += 1

                cplay = [[sg.Text('Asigna categorias y atributos a este producto')],
                          chklist, chklist2,
                          [sg.Ok()]]

                wincp = sg.Window('Nuevo producto', cplay)
                ec, vc = wincp.read()
                wincp.close()

                pcats = [{'id': i} for i in vc if vc[i] == True]
                patri = [{'id': i-1000} for i in vc if vc[i] == True and i > 1000]

                data = {
                "name": v['name'],
                "type": "variable",
                "description": v['descr'],
                "categories": pcats,
                "attributes": patri
                }

                wcapi.post("products", data).json()
                sg.popup_no_wait('Producto creado con exito', non_blocking = True)

    except:
        sg.popup_error('Ha ocurrido un error, ejecute el programa nuevamente')
        return



woo_prods()
import PySimpleGUI as sg
import xlsxwriter
import acciones
import propiedades

cats = propiedades.categorias()
atribs = propiedades.atributos()

sg.theme("DarkGrey3")
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
                acciones.actualizar()
            except:
                sg.popup_error('Es necesario ejecutar "Extraer" antes de actualizar')
                return

            sg.popup_no_wait('Se han actualizado todos los productos', non_blocking = True)

        elif event == 'Extraer':
            sg.popup_no_wait('Extrayendo productos...', non_blocking = True)

            acciones.extraer()

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

                acciones.crear(data)
                
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

                acciones.crear(data)

                sg.popup_no_wait('Producto creado con exito', non_blocking = True)

    except:
        sg.popup_error('Ha ocurrido un error, ejecute el programa nuevamente')
        return



woo_prods()
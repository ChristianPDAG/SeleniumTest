from sel import cargar_excepciones, verificar_excepcion
from inERP import *
import time
import xlwings as xw

app = xw.App(visible=False)
wb = app.books.open('ArchivoDatosAutomatizacion.xlsx')
#wb = xw.Book('ArchivoDatos.xlsx')

sheet = wb.sheets['Hoja1']
pagina = sheet['B1'].value
navegador = sheet['C1'].value
username = sheet['B2'].value
password = sheet['B3'].value
codTipoOperacion = sheet['B4'].value
rut = sheet['B5'].value
descripcion = sheet['B6'].value
codDireccion = sheet['B7'].value
codProducto = sheet['B8'].value
cantProducto = sheet['B9'].value
codBodega = sheet['B10'].value
pUnitarioProducto = sheet['B11'].value
descripcionProducto = sheet['B12'].value
codServicio = sheet['B13'].value
codConsumo = sheet['B14'].value
cantServicio = sheet['B15'].value
pUnitarioServicio = sheet['B16'].value
descripcionServicio = sheet['B17'].value
recepcionar = sheet['B18'].value
descripcionRecepcion = sheet['B19'].value

if navegador == "Chrome":
    driver = webdriver.Chrome()
elif navegador == "Edge":
    driver = webdriver.Edge()
# Referencia a clases
inBot = ordenesCompra(driver)

# Funciones de cambio de frame (el frame es el marco donde se encuentra el DOM)
def switchIframe(by,iframe):
    WebDriverWait(driver,20).until(EC.frame_to_be_available_and_switch_to_it((by,iframe)))
    print(f'Ha ingresado al iframe {iframe}')
    
       
def switchDefaultContent():
    driver.switch_to.default_content()
    print(f"Saliendo del iframe")

driver.maximize_window()
# Ingresar desde el Login / Bucle para determinar que haya ingresado al usuario correctamente
while True:
    try:
        driver.get(pagina)
        assert "Sonda" in driver.title
        switchIframe(By.XPATH,'//*[@id="iframe_body"]')  
        print("La página está cargando")  
        inBot.iptFrame('//*[@id="ctl11_TxtIdPersona"]', username)
        print("Ha ingresado el usuario")
        inBot.iptFrame('//*[@id="ctl11_TxtClave"]', password)
        print("Ha ingresado la password")
        """time.sleep(3)
        inBot.presionarBoton('//*[@id="ctl11_Btnsiguiente"]')
        time.sleep(3)"""
        switchDefaultContent()
        # Operaciones para viajar por el Menu
        inBot.presionarBoton('//*[@id="k_panelbar"]/li[2]/span/span')
        if inBot.sw==1:
            print("Ha ingresado con éxito a su usuario")
            break
    except:
        print("Se ha detectado un ERROR, la página volverá a cargar")

inBot.cerrarVentanaEmergente('//*[@id="iframe_emergente_contenedor"]/span')
inBot.presionarBoton('//*[@id="k_panelbar"]/li[2]/ul/li[17]/span/span')
time.sleep(1) # se le ha dado tiempo para que espere que cargue
inBot.presionarBoton('//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/span/span')
#For para que haga una orden por cada hoja que exista en el Excel

cantHojas = wb.sheets.count
for i in range(cantHojas):
    try:
        #print(i)
        sheet = wb.sheets[i]
        switchDefaultContent()
        print(f'Se está creando la orden de compra de el archivo {sheet}')
        #Aquí comienza el bucle para la creación de una orden
        time.sleep(1)
        inBot.presionarBoton('//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[1]/a')
        time.sleep(3)
        inBot.presionarBoton('//*[@id="btn-toggle-menu"]') # Cierra Menu
        switchIframe(By.NAME,'iframe_body')
        # Pantalla \Orden de Compra
        print("Ingreso orden de compra")
        inBot.presionarBoton('//*[@id="ctl11_btnNuevo"]')
        #AÑADIR SELECCION DE IMPORTADO
        inBot.sobreescribirInput('//*[@id="ctl11_TipoOpeDes_TipoOpeDes_texto"]',codTipoOperacion)
        inBot.iptFrame('//*[@id="ctl11_EntRut"]',rut)
        inBot.presionarBoton('//*[@id="ctl11_hylContactos"]/img')
        time.sleep(3)

        switchDefaultContent()
        switchIframe(By.CLASS_NAME, "k-content-frame")
        # Pantalla de Contactos
        print("Seleccionando contactos")
        inBot.presionarBoton('//*[@id="ctl11_TabContactos"]/tbody/tr[2]/td[1]/a/img')

        switchDefaultContent()
        switchIframe(By.NAME,"iframe_body")
        # Pantalla \Orden de Compra
        inBot.iptFrame('//*[@id="ctl11_AdqGlosa"]',descripcion)
        inBot.presionarBoton('//*[@id="ctl11_btnGuardar"]') #Botón Guardar
        time.sleep(3)

        switchDefaultContent()
        switchIframe(By.CLASS_NAME, "k-content-frame")
        # Pantalla al "Guardar" ingresar direcci+on
        print("Datos comercial orden de compra")
        inBot.iptFrame('//*[@id="ctl11_EntDirDireccion_EntDirDireccion_texto"]',codDireccion)
        inBot.presionarBoton('//*[@id="ctl11_BtnGuardar"]')
        inBot.presionarBoton('//*[@id="ctl11_BtnVolver"]')
        time.sleep(3)

        switchDefaultContent()
        switchIframe(By.NAME,'iframe_body')
        # Pantalla \Orden de Compra -> Botón "Detalle"
        inBot.presionarBoton('//*[@id="Table81"]/tbody/tr/td[2]')
        time.sleep(3)
        ## Desde aquí se debe automatizar productos/servicios ingresados en el excel
        switchDefaultContent()
        switchIframe(By.CLASS_NAME,"k-content-frame")
        # Pantalla \Detalle -> Ingreso producto
        print("Ingreso detalle orden de compra")
        if sheet['B8'].value is not None:
            range1 = sheet['B8'].expand('right')
            for i in range1:
                print("Ingresando producto...")
                letra = i.address.split('$')[1]
                codProducto = sheet[f'{letra}8'].value
                cantProducto = sheet[f'{letra}9'].value
                codBodega = sheet[f'{letra}10'].value
                pUnitarioProducto = sheet[f'{letra}11'].value
                descripcionProducto = sheet[f'{letra}12'].value
                inBot.presionarBoton('//*[@id="ctl11_btnNuevo"]')
                inBot.iptFrame('//*[@id="ctl11_CodigoAlternativo"]',codProducto)
                inBot.iptFrame('//*[@id="ctl11_CantidadCompra"]',cantProducto)
                inBot.presionarBoton('//*[@id="ctl11_BodegaCod"]') #casilla
                inBot.sobreescribirInput('//*[@id="ctl11_BodegaCod"]',codBodega)
                inBot.presionarBoton('//*[@id="ctl11_PrecioUni"]') #casilla
                inBot.iptFrame('//*[@id="ctl11_PrecioUni"]',pUnitarioProducto)
                inBot.presionarBoton('//*[@id="ctl11_AdqGlosa"]') #casilla
                inBot.iptFrame('//*[@id="ctl11_AdqGlosa"]',descripcionProducto)
                print("Guardando producto...")
                inBot.presionarBoton('//*[@id="ctl11_BtnGuardar"]')
                time.sleep(2)

            if sheet['B13'].value is not None:    #Nuevo Servicio
                range2 = sheet['B13'].expand('right')
                for i in range2:
                    print("Ingresando servicio...")
                    letraS = i.address.split('$')[1]
                    codServicio = sheet[f'{letraS}13'].value
                    codConsumo = sheet[f'{letraS}14'].value
                    cantServicio = sheet[f'{letraS}15'].value
                    pUnitarioServicio = sheet[f'{letraS}16'].value
                    descripcionServicio = sheet[f'{letraS}17'].value
                    inBot.presionarBoton('//*[@id="ctl11_btnNuevo"]')
                    inBot.presionarBoton('//*[@id="ctl11_OptTipoItem_1"]')
                    inBot.presionarBoton('//*[@id="ctl11_ServicioCod"]')
                    inBot.iptFrame('//*[@id="ctl11_ServicioCod"]',codServicio)
                    inBot.presionarBoton('//*[@id="ctl11_CConsumoCod"]')
                    inBot.sobreescribirInput('//*[@id="ctl11_CConsumoCod"]',codConsumo)
                    inBot.presionarBoton('//*[@id="ctl11_CantidadCompra"]')
                    inBot.iptFrame('//*[@id="ctl11_CantidadCompra"]',cantServicio)
                    inBot.presionarBoton('//*[@id="ctl11_PrecioUni"]')
                    inBot.iptFrame('//*[@id="ctl11_PrecioUni"]',pUnitarioServicio)
                    inBot.presionarBoton('//*[@id="ctl11_AdqGlosa"]')
                    inBot.iptFrame('//*[@id="ctl11_AdqGlosa"]',descripcionServicio)
                    print("Guardando Servicio...")
                    inBot.presionarBoton('//*[@id="ctl11_BtnGuardar"]')
                    #inBot.presionarBoton('//*[@id="ctl11_BtnVolver"]')
                    time.sleep(1)
            else: 
                print("No hay servicios para agregar")

        elif sheet['B13'].value is not None:    #Nuevo Servicio
            range2 = sheet['B13'].expand('right')
            for i in range2:
                print("Ingresando servicio")
                letraS = i.address.split('$')[1]
                codServicio = sheet[f'{letraS}13'].value
                codConsumo = sheet[f'{letraS}14'].value
                cantServicio = sheet[f'{letraS}15'].value
                pUnitarioServicio = sheet[f'{letraS}16'].value
                descripcionServicio = sheet[f'{letraS}17'].value
                inBot.presionarBoton('//*[@id="ctl11_btnNuevo"]')
                inBot.presionarBoton('//*[@id="ctl11_OptTipoItem_1"]')
                inBot.presionarBoton('//*[@id="ctl11_ServicioCod"]')
                inBot.iptFrame('//*[@id="ctl11_ServicioCod"]',codServicio)
                inBot.presionarBoton('//*[@id="ctl11_CConsumoCod"]')
                inBot.sobreescribirInput('//*[@id="ctl11_CConsumoCod"]',codConsumo)
                inBot.presionarBoton('//*[@id="ctl11_CantidadCompra"]')
                inBot.iptFrame('//*[@id="ctl11_CantidadCompra"]',cantServicio)
                inBot.presionarBoton('//*[@id="ctl11_BodegaCod"]')
                inBot.sobreescribirInput('//*[@id="ctl11_BodegaCod"]',codBodega)
                inBot.presionarBoton('//*[@id="ctl11_PrecioUni"]')
                inBot.iptFrame('//*[@id="ctl11_PrecioUni"]',pUnitarioServicio)
                inBot.presionarBoton('//*[@id="ctl11_AdqGlosa"]')
                inBot.iptFrame('//*[@id="ctl11_AdqGlosa"]',descripcionServicio)
                print("Guardando Servicio...")
                inBot.presionarBoton('//*[@id="ctl11_BtnGuardar"]')
                #inBot.presionarBoton('//*[@id="ctl11_BtnVolver"]')
        else:
            print('No se han ingresado productos ni servicios para el Detalle')

        inBot.presionarBoton('//*[@id="ctl11_BtnVolver"]')
        time.sleep(3)
        #Volviendo al ingreso de orden
        
        switchDefaultContent()
        switchIframe(By.NAME,'iframe_body')
        elemento = driver.find_element(By.XPATH,'//*[@id="ctl11_CorrLogtNum"]')
        element = driver.find_element(By.XPATH,'//*[@id="ctl11_OcCabId"]')
        numeroOrden = element.get_attribute('value')
        valorCorrelativo = elemento.get_attribute('value')
        sheet['B21'].value = numeroOrden
        sheet['B22'].value = valorCorrelativo # Almacena valor correlativo de la orden
        print(f'Aprobando Orden de Compra N°: {numeroOrden} \nCorrelativo: {valorCorrelativo} ')
        inBot.btnOpcionAprobar('//*[@id="lnk12"]')
        time.sleep(3)
        #Ingresar al menú
        switchDefaultContent()
        inBot.presionarBoton('//*[@id="btn-toggle-menu"]')
        time.sleep(1)
        inBot.presionarBoton('//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[2]/a')
        time.sleep(1)
        if recepcionar == 'S' or recepcionar == 's':
            time.sleep(3)
            #Recepción de orden
            switchIframe(By.NAME,'iframe_body')
            inBot.presionarBoton('//*[@id="ctl11_btnNuevo"]')
            print("Ingresando recepción ordenes de compra")
            inBot.iptFrame('//*[@id="ctl11_BodegaCod"]',codBodega)
            inBot.presionarBoton('//*[@id="ctl11_GlosaExis"]')
            inBot.iptFrame('//*[@id="ctl11_GlosaExis"]',descripcionRecepcion)


            inBot.presionarBoton('//*[@id="ctl11_CorrLogtNum"]')
            inBot.iptFrame('//*[@id="ctl11_CorrLogtNum"]',valorCorrelativo)
            inBot.presionarBoton('//*[@id="ctl11_GlosaExis"]')
            inBot.presionarBoton('//*[@id="ctl11_btnGuardar"]')

            time.sleep(3)
            inBot.btnOpcionAprobar('//*[@id="lnk7"]')
            time.sleep(2)
            time.sleep(1)
            print(f'La orden de compra N° {numeroOrden} ha sido recepcionada con éxito')
        else:
            print(f'No se ha recepcionado la Orden de Compra {numeroOrden}')
    except:
        driver.switch_to.default_content()
        switchIframe(By.CLASS_NAME, "k-content-frame")
        archivo_xml = 'SondaExceptions.xml'  # Reemplaza con tu ruta real
        excepciones = cargar_excepciones(archivo_xml)
        print("Está verificando excepciones")
        verificar_excepcion(driver, excepciones)
        inBot.cerrarVentanaEmergente('/html/body/div[5]/div[1]/div/a')
        inBot.presionarBoton('//*[@id="btn-toggle-menu"]')
    #switchDefaultContent()
    #inBot.presionarBoton('//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[1]/a')
    

# Mantiene el navegador abierto
"""if navegador == "Chrome":
    options = webdriver.ChromeOptions()
elif navegador == "Edge":
    options = webdriver.EdgeOptions()"""
wb.save()
wb.close()
app.quit()

#driver = webdriver.Chrome(options=options)
input("Presiona Enter para detener el bot... \nADVERTENCIA: AL DETENER EL BOT TAMBIEN SE CIERRA LA PESTAÑA ACTUAL, \nSE RECOMIENDA SOLO DETENER CUANDO YA HAZ TERMINADO EL USO DE ESTA PESTAÑA")
driver.quit()
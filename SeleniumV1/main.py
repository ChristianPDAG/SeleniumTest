from selenium import webdriver
import time
import sys
from operaciones_erp import *
import xlwings as xw
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException, StaleElementReferenceException, ElementClickInterceptedException


class TestOrden:
    iframe_body_name = 'iframe_body'
    iframe_class_name = "k-content-frame"
    text_box_username = '//*[@id="ctl11_TxtIdPersona"]'
    text_box_password = '//*[@id="ctl11_TxtClave"]'
    text_box_tipo_operacion = '//*[@id="ctl11_TipoOpeDes_TipoOpeDes_texto"]'
    text_box_rut = '//*[@id="ctl11_EntRut"]'
    text_box_descripcion = '//*[@id="ctl11_AdqGlosa"]'
    text_box_cod_direccion = '//*[@id="ctl11_EntDirDireccion_EntDirDireccion_texto"]'
    text_box_cod_producto = '//*[@id="ctl11_CodigoAlternativo"]'
    text_box_cantidad = '//*[@id="ctl11_CantidadCompra"]'
    text_box_cod_bodega = '//*[@id="ctl11_BodegaCod"]'
    text_box_precio_unitario = '//*[@id="ctl11_PrecioUni"]'
    text_box_cod_servicio = '//*[@id="ctl11_ServicioCod"]'
    text_box_cod_consumo = '//*[@id="ctl11_CConsumoCod"]'
    text_box_correlativo = '//*[@id="ctl11_CorrLogtNum"]'
    text_box_num_orden = '//*[@id="ctl11_OcCabId"]'
    text_box_descripcion_recepcion = '//*[@id="ctl11_GlosaExis"]'
    btn_inicio_sesion = '//*[@id="ctl11_Btnsiguiente"]'
    btn_transacciones_menu = '//*[@id="k_panelbar"]/li[2]/span/span'
    btn_X_ventana_emergente = '//*[@id="iframe_emergente_contenedor"]/span'
    btn_adquisiciones_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/span/span'
    btn_administracion_compra_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/span/span'
    btn_orden_compra_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[1]/a'
    btn_abrir_cerrar_menu = '//*[@id="btn-toggle-menu"]'
    btn_nuevo = '//*[@id="ctl11_btnNuevo"]'
    btn_buscar_contactos = '//*[@id="ctl11_hylContactos"]/img'
    btn_seleccionar_contacto = '//*[@id="ctl11_TabContactos"]/tbody/tr[2]/td[1]/a/img'
    btn_guardar = '//*[@id="ctl11_btnGuardar"]'
    btn_guardar_mayuscula = '//*[@id="ctl11_BtnGuardar"]'
    btn_volver = '//*[@id="ctl11_BtnVolver"]'
    btn_detalle = '//*[@id="Table81"]/tbody/tr/td[2]'
    btn_aprobar_orden = '//*[@id="lnk12"]'
    btn_aprobar_recepcion = '//*[@id="lnk7"]'
    btn_recepcion_orden_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[2]/a'
    radio_servicio = '//*[@id="ctl11_OptTipoItem_1"]'

    def __init__(self,archivo_xlsx):
        self.app = xw.App(visible=False)
        self.archivo_xlsx = archivo_xlsx
        self.workbook = self.app.books.open(self.archivo_xlsx)
        self.hoja = self.workbook.sheets['CabeceraOC']
        self.hoja2 = self.workbook.sheets['DetalleOC']
 
        navegador = self.hoja['C2'].value
        if navegador == "Chrome":
            self.driver = webdriver.Chrome()
        elif navegador == "Edge":
            self.driver = webdriver.Edge()
        self.action_bot = OperacionesOC(self.driver)
        self.driver.maximize_window()

    def test_orden(self):    
        username = self.hoja['A2'].value
        password = self.hoja['B2'].value
        pagina = self.hoja['D2'].value
        if pagina == "https://garza1.sonda.com/fin700QACentral75":
            print(f"Ingresando a '{pagina}'")
        else:
            raise Exception("Pagina ingresada es invalida para esta automatizacion\nPor favor ingresar pagina valida ")
        cc=0
        while True:
            self.driver.get(pagina)
            assert "Sonda" in self.driver.title
            self.action_bot.switchIframe(By.NAME, self.iframe_body_name)  
            print("La pagina esta cargando")  
            self.action_bot.iptFrame(self.text_box_username, username)
            self.action_bot.iptFrame(self.text_box_password , password)
            #self.action_bot.presionarBoton(self.btn_inicio_sesion)
            self.action_bot.switchDefaultContent()
            # Operaciones para viajar por el Menu
            try:
                self.action_bot.presionarBoton(self.btn_transacciones_menu)
            except Exception as e:
                print("Se ha detectado un ERROR  ")
            if self.action_bot.sw==1:
                print("Ha ingresado con exito a su usuario")
                break
            else:
                print("La pagina volvera a cargar...")
                cc+=1
                if cc == 2:
                    raise Exception("Se ha detectado un usuario o clave incorrecta\nPor favor revise su hoja de calculo 'ArchivoDatosAutomatizacionOrdenDeCompra.xlsx' y vuelva a intentar\nEl programa esta finalizando...")
                
        self.action_bot.cerrarVentanaEmergente(self.btn_X_ventana_emergente)
        self.action_bot.presionarBoton(self.btn_adquisiciones_menu)
        time.sleep(1) # se le ha dado tiempo para que espere que cargue
        self.action_bot.presionarBoton(self.btn_administracion_compra_menu)
        data = self.hoja.range((4,1),(4,7)).value
        data2 = self.hoja2['A2:J2'].expand('down').rows
        if data is not None:
           range1 = self.hoja['A4:H4'].expand('down').rows
           for i in range1:      
                IDOC_Cabecera = str(int(i[0].value))
                codTipoOperacion = str(int(i[1].value))
                rut = i[2].value
                descripcion = i[3].value
                if descripcion is None:
                    descripcion = "Descripcion"
                codDireccion = str(int(i[4].value))
                descripcionRecepcion = i[5].value
                if descripcionRecepcion is None:
                    descripcionRecepcion = "Recepcion"
                self.action_bot.switchDefaultContent()
                print(f'Se esta creando la orden de compra con IDOC {IDOC_Cabecera}')
                #Aqu� comienza el bucle para la creacion de una orden
                time.sleep(1)
                self.action_bot.presionarBoton(self.btn_orden_compra_menu)
                time.sleep(3)
                self.action_bot.presionarBoton(self.btn_abrir_cerrar_menu) # Cierra Menu
                self.action_bot.switchIframe(By.NAME,self.iframe_body_name)
                # Pantalla \Orden de Compra
                print("Ingreso orden de compra")
                self.action_bot.presionarBoton(self.btn_nuevo)
                if codTipoOperacion == '85':
                    self.action_bot.sobreescribirInput(self.text_box_tipo_operacion, codTipoOperacion)
                    print(f"Codigo tipo de operacion '{codTipoOperacion}' ingresado")
                else:
                    raise Exception("ATENCION: Codigo de tipo operacion invalido\nLa version del bot esta hecha solo para el codigo '85' que hace referencia a 'Compra Nacional'\n\nPor favor ingrese un codigo tipo operacion valido en su hoja de calculo 'ArchivoDatosAutomatizacionOrdenDeCompra.xlsx' y vuelva a intentar")
                self.action_bot.iptFrame(self.text_box_rut,rut)
                print(f"Rut '{rut}' ingresado")
                try:
                    self.action_bot.presionarBoton(self.btn_buscar_contactos)
                except ElementClickInterceptedException as e:
                    print("ATENCION: Rut invalido\nPor favor ingrese un rut valido en su hoja de calculo 'ArchivoDatosAutomatizacionOrdenDeCompra.xlsx' y vuelva a intentar")
                time.sleep(3)
                self.action_bot.switchDefaultContent()
                self.action_bot.switchIframe(By.CLASS_NAME, self.iframe_class_name)
                # Pantalla de Contactos
                self.action_bot.presionarBoton(self.btn_seleccionar_contacto)
                print("Seleccionando contactos")
                self.action_bot.switchDefaultContent()
                self.action_bot.switchIframe(By.NAME,self.iframe_body_name)
                # Pantalla \Orden de Compra
                self.action_bot.iptFrame(self.text_box_descripcion, descripcion)
                print(f"Descripcion de Orden '{descripcion}' ingresada")
                self.action_bot.presionarBoton(self.btn_guardar) 
                time.sleep(3)
                self.action_bot.switchDefaultContent()
                self.action_bot.switchIframe(By.CLASS_NAME, self.iframe_class_name)
                # Pantalla al "Guardar" ingresar direcci+on
                print("Datos comercial orden de compra")
                if codDireccion in ["2397","2377","2408","2376","2385","2379","2382","2388","2390","2389","2407","220","2391","2409","223","2401","2400","2398","2395","2393","2399","2378","2383","2394","2386","2380","2392","2384","2403","2402","2404","4","2387","2406","2396","2405","2381"]:
                    self.action_bot.iptFrame(self.text_box_cod_direccion, codDireccion)
                    print(f"Codigo de direccion '{codDireccion}' ingresado")
                    self.action_bot.presionarBoton(self.btn_guardar_mayuscula)
                    print("Guardando datos comercial")
                else:
                    raise Exception("ATENCION: Codigo de direccion invalido\nPor favor ingrese un codigo de direccion valido en su hoja de calculo 'ArchivoDatosAutomatizacionOrdenDeCompra.xlsx' y vuelva a intentar")
                time.sleep(5)
                self.action_bot.presionarBoton(self.btn_volver)
                time.sleep(3)
                self.action_bot.switchDefaultContent()
                self.action_bot.switchIframe(By.NAME,self.iframe_body_name)
                # Pantalla \Orden de Compra -> Boton "Detalle"
                self.action_bot.presionarBoton(self.btn_detalle)
                time.sleep(3)
                ## Desde aqu� se debe automatizar productos/servicios ingresados en el excel
                self.action_bot.switchDefaultContent()
                self.action_bot.switchIframe(By.CLASS_NAME,self.iframe_class_name)
                for e in data2: #Datos en Detalle OC
                    IDOC_Detalle = str(int(e[0].value))
                    codProducto = e[1].value
                    codServicio = e[2].value
                    codBodega = str(int(e[3].value))
                    cantidad = str(int(e[5].value))
                    pUnitario = str(int(e[6].value))
                    descripcionDetalle = e[7].value
                    if descripcionDetalle is None:
                        descripcionDetalle = "Detalle"
                    if IDOC_Detalle == IDOC_Cabecera:
                        print("Ingreso detalle orden de compra")
                        if codProducto is not None:
                            print("Registrando producto...")
                            self.action_bot.presionarBoton(self.btn_nuevo)
                            self.action_bot.iptFrame(self.text_box_cod_producto, codProducto)
                            print(f"Codigo de producto '{codProducto}' ingresado")
                            self.action_bot.presionarBoton(self.text_box_cantidad)
                            print("Cantidad presionada")
                            self.action_bot.iptFrame(self.text_box_cantidad, cantidad)
                            print(f"Cantidad '{cantidad}' ingresada")
                            self.action_bot.presionarBoton(self.text_box_cod_bodega) #casilla
                            self.action_bot.sobreescribirInput(self.text_box_cod_bodega, codBodega)
                            print(f"Codigo de bodega '{codBodega}' ingresado")
                            self.action_bot.presionarBoton(self.text_box_precio_unitario) #casilla
                            self.action_bot.iptFrame(self.text_box_precio_unitario, pUnitario)
                            print(f"Precio unitario '{pUnitario}' ingresado")
                            self.action_bot.presionarBoton(self.text_box_descripcion) #casilla
                            self.action_bot.iptFrame(self.text_box_descripcion, descripcionDetalle)
                            print(f"Descripción de producto '{descripcionDetalle}' ingresado")
                            print("Guardando producto...")
                            self.action_bot.presionarBoton(self.btn_guardar_mayuscula)
                            print(f"Producto {codProducto} registrado exitosamente !")
                            time.sleep(2)
                        elif codServicio is not None:
                            codConsumo = str(int(e[4].value))
                            print("Registrando servicio...")
                            self.action_bot.presionarBoton(self.btn_nuevo)
                            self.action_bot.presionarBoton(self.radio_servicio)
                            self.action_bot.presionarBoton(self.text_box_cod_servicio)
                            self.action_bot.iptFrame(self.text_box_cod_servicio, codServicio)
                            self.action_bot.presionarBoton(self.text_box_cod_consumo)
                            print(f"Codigo de servicio '{codServicio}' ingresado")
                            self.action_bot.sobreescribirInput(self.text_box_cod_consumo, codConsumo)
                            self.action_bot.presionarBoton(self.text_box_cantidad)
                            print(f"Codigo de consumo '{codConsumo}' ingresado ")
                            self.action_bot.iptFrame(self.text_box_cantidad, cantidad)
                            self.action_bot.presionarBoton(self.text_box_cod_bodega) #casilla
                            print(f"Cantidad '{cantidad}' ingresada")
                            self.action_bot.sobreescribirInput(self.text_box_cod_bodega, codBodega)
                            self.action_bot.presionarBoton(self.text_box_precio_unitario)
                            print(f"Codigo de bodega '{codBodega}' ingresado")
                            self.action_bot.iptFrame(self.text_box_precio_unitario, pUnitario)
                            self.action_bot.presionarBoton(self.text_box_descripcion)
                            print(f"Precio unitario '{pUnitario}' ingresado")
                            self.action_bot.iptFrame(self.text_box_descripcion, descripcion)
                            print(f"Descripcion de servicio '{descripcion}' ingresado")
                            print("Guardando Servicio...")
                            self.action_bot.presionarBoton(self.btn_guardar_mayuscula)
                            print(f"Servicio {codServicio} registrado exitosamente !")

                            #self.action_bot.presionarBoton('//*[@id="ctl11_BtnVolver"]')
                            time.sleep(1)
                        else: 
                            print("No hay más productos y/o servicios para agregar")

                       
                self.action_bot.presionarBoton(self.btn_volver)
                time.sleep(3)
                #Volviendo al ingreso de orden
        
                self.action_bot.switchDefaultContent()
                self.action_bot.switchIframe(By.NAME,self.iframe_body_name)
                elemento = self.driver.find_element(By.XPATH, self.text_box_correlativo)
                element = self.driver.find_element(By.XPATH, self.text_box_num_orden)
                numeroOrden = element.get_attribute('value')
                valorCorrelativo = elemento.get_attribute('value')
                i[6].value = numeroOrden
                i[7].value = valorCorrelativo # Almacena valor correlativo de la orden
                print(f'Aprobando Orden de Compra \nNumero: {numeroOrden} \nCorrelativo: {valorCorrelativo} ')
                self.action_bot.btnOpcionAprobar(self.btn_aprobar_orden)
                time.sleep(3)
                #Ingresar al men�
                self.action_bot.switchDefaultContent()
                self.action_bot.presionarBoton(self.btn_abrir_cerrar_menu)
                time.sleep(1)
                self.action_bot.presionarBoton(self.btn_recepcion_orden_menu)
                time.sleep(2)            
                #Recepcion de orden
                self.action_bot.switchIframe(By.NAME, self.iframe_body_name)
                self.action_bot.presionarBoton(self.btn_nuevo)
                print("Ingresando recepcion ordenes de compra")
                self.action_bot.iptFrame(self.text_box_cod_bodega, codBodega)
                self.action_bot.presionarBoton(self.text_box_descripcion_recepcion)
                print(f"Codigo de bodega '{codBodega}' ingresado")
                self.action_bot.iptFrame(self.text_box_descripcion_recepcion, descripcionRecepcion)
                print(f"Descripcion de recepcion '{descripcionRecepcion}' ingresado")
                self.action_bot.presionarBoton(self.text_box_correlativo)
                self.action_bot.iptFrame(self.text_box_correlativo, valorCorrelativo)
                self.action_bot.presionarBoton(self.text_box_descripcion_recepcion)
                self.action_bot.presionarBoton(self.btn_guardar)
                time.sleep(1)
                self.action_bot.btnOpcionAprobar(self.btn_aprobar_recepcion)
                time.sleep(1)
                print(f'La orden de compra Numero {numeroOrden} ha sido recepcionada con exito')


    def teardown(self):
        self.workbook.save()
        self.workbook.close()
        self.app.quit()
        input("Presiona Enter para detener el bot... \nADVERTENCIA: AL DETENER EL BOT TAMBIEN SE CIERRA LA PESTANA ACTUAL, \nSE RECOMIENDA SOLO DETENER CUANDO YA HAZ TERMINADO EL USO DE ESTA PESTANA")
        self.driver.quit()
        sys.exit()


if __name__ == '__main__':
    init_bot = TestOrden(archivo_xlsx = 'ArchivoDatosAutomatizacionOrdenDeCompra.xlsx')
    try:
        init_bot.test_orden()
    except ElementClickInterceptedException as e:
        print(e)
    except ElementClickInterceptedException as ep:
        print(ep)
    except Exception as e:
        print(e)
    finally:
        init_bot.teardown()

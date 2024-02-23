from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton
from selenium.common.exceptions import NoSuchElementException, TimeoutException


class ordenesCompra:
    def __init__(self,driver):
        self.driver = driver
        self.sw = 0

    def cerrarVentanaEmergente(self,nameInDOMXpath): # Cierra ventana emergente que aparece al ingresar al usuario.
        self.driver.switch_to.default_content() 
        wait = WebDriverWait(self.driver, timeout=20)
        wait.until(EC.presence_of_element_located((By.XPATH, nameInDOMXpath)))
        scriptX = self.driver.find_element(By.XPATH,nameInDOMXpath)
        self.driver.execute_script("arguments[0].click();",scriptX)
        print("Se ha cerrado la ventana emergente (X)")
        
            

    def presionarBoton(self,nameInDOMXpath): # Presionar botón
        wait = WebDriverWait(self.driver, timeout=20)
        wait.until(EC.visibility_of_element_located((By.XPATH, nameInDOMXpath)))
        self.driver.find_element(By.XPATH,nameInDOMXpath).click()
        self.sw=1 
        

    def forzarValor(self,nameInDOMXpath,valorOperacion): # Se usa para forzar el cambio de "value" de un elemento // Solo en casos excepcionales
        tipoOp = self.driver.find_element(By.XPATH,nameInDOMXpath)
        self.driver.execute_script(f"arguments[0].value = {valorOperacion};",tipoOp).click()
        wait = WebDriverWait(self.driver, timeout=5)
        wait.until(EC.text_to_be_present_in_element_value((By.XPATH, nameInDOMXpath),valorOperacion))
        print(f"Se ha forzado el ingreso del valor {valorOperacion}")

    def iptFrame(self,nameInDOMXpath,valor): # Ingreso de datos en un <input>
        wait = WebDriverWait(self.driver, timeout=20)
        wait.until(EC.presence_of_element_located((By.XPATH, nameInDOMXpath)))
        campo = self.driver.find_element(By.XPATH,nameInDOMXpath)
        self.borrarEscribir(campo,valor)
        campo.send_keys(Keys.ENTER)



    def sobreescribirInput(self,nameInDOMXpath,valor): # Esta función ha sido creada para escribir en las casillas que ya traen un valor por defecto
        wait = WebDriverWait(self.driver,timeout=5)
        wait.until(EC.presence_of_element_located((By.XPATH,nameInDOMXpath)))
        elemento = self.driver.find_element(By.XPATH,nameInDOMXpath)
        valorInicial = elemento.get_attribute('value')
        elemento.send_keys(Keys.CONTROL + "a")
        elemento.send_keys(valor)
        elemento.send_keys(Keys.ENTER)
        print("Sobreescrito correctamente")
        

    def btnOpcionAprobar(self,nameDivXpath):
        div = self.driver.find_element(By.XPATH, nameDivXpath)
        actions = ActionChains(self.driver)
        actions.move_to_element(div).perform()    
        opcion = self.driver.find_element(By.XPATH,'//a[contains(text(),"Aprobar")]')
        opcion.click()
        print('Orden de compra Aprobada')
                

    """
    while True:
            texto = input('¿Que opción desea presionar? (Aprobar/Anular) \n> ').capitalize()
            if texto == 'Aprobar':
                opcion = self.driver.find_element(By.XPATH,f'//a[contains(text(),"{texto}")]')
                opcion.click()
                print('Orden de compra Aprobada')
                break
            elif texto == 'Anular':
                opcion = self.driver.find_element(By.XPATH,f'//a[contains(text(),"{texto}")]')
                opcion.click()
                print('Se ha Anulado su orden de compra ')
                input('Presione Enter para cerrar el Bot...')
                self.driver.quit()
                break
            else: 
                input('Ingrese una opción válida (Aprobar/Anular)...\nEnter para volver a ingresar')
    """

    def borrarEscribir(self,campo,texto):
        campo.clear()
        campo.send_keys(texto)

    iframe_body_name = 'iframe_body'
    iframe_class_name = "k-content-frame"
    text_box_username = '//*[@id="ctl11_TxtIdPersona"]'
    text_box_password = '//*[@id="ctl11_TxtClave"]'
    text_box_tipo_operacion = '//*[@id="ctl11_TipoOpeDes_TipoOpeDes_texto"]'
    text_box_rut = '//*[@id="ctl11_EntRut"]'
    text_box_descripcion = '//*[@id="ctl11_AdqGlosa"]'
    text_box_cod_direccion = '//*[@id="ctl11_EntDirDireccion_EntDirDireccion_texto"]'
    btn_inicio_sesion = '//*[@id="ctl11_Btnsiguiente"]'
    btn_transacciones_menu = '//*[@id="k_panelbar"]/li[2]/span/span'
    btn_X_ventana_emergente = '//*[@id="iframe_emergente_contenedor"]/span'
    btn_transacciones_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/span/span'
    btn_administracion_compra_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/span/span'
    btn_orden_compra_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[1]/a'
    btn_abrir_cerrar_menu = '//*[@id="btn-toggle-menu"]'
    btn_nueva_orden = '//*[@id="ctl11_btnNuevo"]'
    btn_buscar_contactos = '//*[@id="ctl11_hylContactos"]/img'
    btn_seleccionar_contacto = '//*[@id="ctl11_TabContactos"]/tbody/tr[2]/td[1]/a/img'
    btn_guardar = '//*[@id="ctl11_btnGuardar"]'
    btn_volver = '//*[@id="ctl11_BtnVolver"]'
    btn_detalle = '//*[@id="Table81"]/tbody/tr/td[2]'
    text_box_cod_producto = '//*[@id="ctl11_CodigoAlternativo"]'
    text_box_cantidad = '//*[@id="ctl11_CantidadCompra"]'
    text_box_cod_bodega = '//*[@id="ctl11_BodegaCod"]'
    text_box_precio_unitario = '//*[@id="ctl11_PrecioUni"]'
    radio_servicio = '//*[@id="ctl11_OptTipoItem_1"]'
    text_box_cod_servicio = '//*[@id="ctl11_ServicioCod"]'
    text_box_cod_consumo = '//*[@id="ctl11_CConsumoCod"]'
    text_box_correlativo = '//*[@id="ctl11_CorrLogtNum"]'
    text_box_num_orden = '//*[@id="ctl11_OcCabId"]'
    btn_aprobar_orden = '//*[@id="lnk12"]'
    btn_recepcion_orden_menu = '//*[@id="k_panelbar"]/li[2]/ul/li[17]/ul/li[4]/ul/li[2]/a'
    text_box_descripcion_recepcion = '//*[@id="ctl11_GlosaExis"]'
    btn_aprobar_recepcion = '//*[@id="lnk7"]'
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


class OperacionesOC:
    def __init__(self,driver):
        self.driver = driver
        self.sw = 0

    def cerrarVentanaEmergente(self,nameInDOMXpath): # Cierra ventana emergente que aparece al ingresar al usuario.
        self.driver.switch_to.default_content() 
        wait = WebDriverWait(self.driver, timeout=10)
        wait.until(EC.presence_of_element_located((By.XPATH, nameInDOMXpath)))
        scriptX = self.driver.find_element(By.XPATH,nameInDOMXpath)
        self.driver.execute_script("arguments[0].click();",scriptX)
        print("Se ha cerrado la ventana emergente (X)")           

    def presionarBoton(self,nameInDOMXpath): # Presionar botón

        wait = WebDriverWait(self.driver, timeout=10)
        wait.until(EC.visibility_of_element_located((By.XPATH, nameInDOMXpath)))
        self.driver.find_element(By.XPATH,nameInDOMXpath).click()
        self.sw=1 


    def forzarValor(self,nameInDOMXpath,valorOperacion): # Se usa para forzar el cambio de "value" de un elemento // Solo en casos excepcionales
        tipoOp = self.driver.find_element(By.XPATH,nameInDOMXpath)
        self.driver.execute_script(f"arguments[0].value = {valorOperacion};",tipoOp).click()
        wait = WebDriverWait(self.driver, timeout=5)
        wait.until(EC.text_to_be_present_in_element_value((By.XPATH, nameInDOMXpath),valorOperacion))

    def iptFrame(self,nameInDOMXpath,valor): # Ingreso de datos en un <input>
        wait = WebDriverWait(self.driver, timeout=10)
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

    def switchIframe(self,by,iframe):
        WebDriverWait(self.driver,20).until(EC.frame_to_be_available_and_switch_to_it((by,iframe)))    
       
    def switchDefaultContent(self):
        self.driver.switch_to.default_content()


from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton
from selenium.common.exceptions import NoSuchElementException, TimeoutException


class OperacionesMDP:
    def __init__(self,driver):
        self.driver = driver

    def presionar_boton(self, name_in_dom_xpath: str): # Presionar botón
        wait = WebDriverWait(self.driver, timeout=90)
        wait.until(EC.element_to_be_clickable((By.XPATH, name_in_dom_xpath)))
        self.driver.find_element(By.XPATH,name_in_dom_xpath).click()

    def ingresar_texto(self, name_in_dom_xpath: str, valor: str): # Ingreso de datos en un <input>
        wait = WebDriverWait(self.driver, timeout=60)
        wait.until(EC.presence_of_element_located((By.XPATH, name_in_dom_xpath)))
        campo = self.driver.find_element(By.XPATH,name_in_dom_xpath)
        self.borrar_escribir(campo,valor)
        #campo.send_keys(Keys.ENTER)
    
    def select_text(self, name_in_dom_xpath: str, valor: str):
        wait = WebDriverWait(self.driver, timeout=60)
        wait.until(EC.presence_of_element_located((By.XPATH, name_in_dom_xpath)))
        select_element = self.driver.find_element(By.XPATH, name_in_dom_xpath)
        select = Select(select_element)
        select.select_by_value(valor)

    def borrar_escribir(self, campo ,texto: str):
        campo.clear()
        campo.send_keys(texto)



    

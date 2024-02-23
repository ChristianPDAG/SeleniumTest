import xml.etree.ElementTree as ET
from selenium import webdriver
import time
import random
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton
from selenium.common.exceptions import NoSuchElementException, TimeoutException


def cargar_excepciones(archivo_xml):
    excepciones = {}
    tree = ET.parse(archivo_xml)
    root = tree.getroot()

    for data_element in root.findall('.//DATA'):
            codigo = data_element.find('CODIGO').text
            descripcion = data_element.find('DESCRIPCION').text
            severidad = data_element.find('SEVERIDAD').text
            proceso_negocio = data_element.find('PROCESONEGOCIO').text

            excepciones[codigo] = {
                'descripcion': descripcion,
                'severidad': severidad,
                'proceso_negocio': proceso_negocio
            }

    return excepciones


def verificar_excepcion(driver,excepciones):
    try: 
        sw=0
        try:
            print("Cargando elementos_mensaje...")
            elemento_mensaje = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="lblErrorCompleto"]')))
            sw=1
            print(sw)
        except:
            
            print("No se encontró un tipo de ventana emergente")

        try:
            print("Cargando otro_mensaje...")
            otro_mensaje = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="theBody"]')))
            driver.find_element(By.XPATH,'//*[@id="theBody"]')
            sw=2
            print(sw)
        except:
            print("No se encontró otro tipo de ventana emergente")

        print("Elemento encontrado:")
    # Obtener el mensaje de la excepción
        if sw==1:
            mensaje_actual = elemento_mensaje.text
            
            print(mensaje_actual)
            print("ENTRANDO AL BUCLE")

        
            for codigo, info in excepciones.items():
                descripcion = info['descripcion']
                severidad = info['severidad']
                proceso_negocio = info['proceso_negocio']
                #print(f"Comparando con código {codigo} y descripción {descripcion}")
                if descripcion and descripcion in mensaje_actual :
                    print(f"Se detectó la excepción con código {codigo}:")
                    print(f"Descripción: {descripcion}\nSeveridad: {severidad}\nProceso de negocio: {proceso_negocio} ")
                    i=0
                    driver.save_screenshot(f'./{codigo};{descripcion};imagen-{i}.png')
                    i+=1
                    time.sleep(1)
                    

        elif sw==2:
            print("Entrando a la tabla")
            otro_mensaje_actual = otro_mensaje.text
            tbody = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl07_table1"]/tbody')))
            print("tbody encontrado")
            filas = tbody.find_elements(By.TAG_NAME, 'tr')  
            errores_encontrados = {}          
            for fila in filas[2:]:
                celdas = fila.find_elements(By.TAG_NAME,'td')
                """for celda in celdas:
                    errores_encontrados = celda.text"""
                if len(celdas) >= 2:
                    for celda in celdas[1:]:
                        print(celda.text)
                        codigo_error = celdas[1].text
                        descripcion_error = celdas[2].text
                        errores_encontrados[codigo_error] = descripcion_error

            print(errores_encontrados)
            num = random.randint(1, 100)
            driver.save_screenshot(f'./;imagenError-{num}.png')
            time.sleep(1)
        else:
            print("No se ha encontrado ningun elemento")

    except NoSuchElementException as e:
        print(f"No se pudo encontrar el elemento: {e}")
    except TimeoutException as e:
        print(f"Tiempo de espera excedido al esperar el elemento: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")


       

archivo_xml = 'SondaExceptions.xml'  # Reemplaza con tu ruta real
excepciones = cargar_excepciones(archivo_xml)

# Inicializar el navegador
"""driver = webdriver.Chrome()
driver.get("https://garza1.sonda.com/fin700QACentral75")
time.sleep(3)
switchIframe(By.NAME,'iframe_body')
except NoSuchElementException as e:
            print(f"No se pudo encontrar el elemento: {e}")
        except TimeoutException as e:
            print(f"Tiempo de espera excedido al esperar el elemento: {e}")
        except Exception as e:
            print(f"Error inesperado: {e}")
"""
# Verificar si aparece alguna excepción en la página
#verificar_excepcion(driver, excepciones)  # Reemplaza con el código específico que esperas

# Continuar con otras acciones
#input("DETENiDO")
# Cerrar el navegador

    # Returns and base64 encoded string into image
#driver.save_screenshot('./image.png')

#driver.quit()
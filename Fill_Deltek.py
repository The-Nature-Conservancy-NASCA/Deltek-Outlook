# Libreria
import numpy as np
from selenium import webdriver
#from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import pandas as pd
import numpy as np
from zipfile import ZipFile
# import numpy as np

# ----------------------------------------------------------------------------------------------------------------------
# Entradas al Sistema
# ----------------------------------------------------------------------------------------------------------------------
LoginID     = ''
Password    = ''
Domain      = 'TNC.ORG'
Path        = r'chromedriver.exe'

# ----------------------------------------------------------------------------------------------------------------------
# Leer Datos
# ----------------------------------------------------------------------------------------------------------------------
Project = pd.read_excel('00-Projects.xlsx')
Datos   = pd.read_csv('02-Deltek.csv')
Value   = pd.merge(Project,Datos, on="Name")

# ----------------------------------------------------------------------------------------------------------------------
# Abrir Google Chrome operado con Selenium cambiando la carpeta de descargas
# ----------------------------------------------------------------------------------------------------------------------
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--start-maximized')
chrome_options.add_argument('--desable-extensions')
driver = webdriver.Chrome(Path, chrome_options=chrome_options)

# ----------------------------------------------------------------------------------------------------------------------
# Abrir la página de Deltek
# ----------------------------------------------------------------------------------------------------------------------
driver.get('https://tnc.hostedaccess.com/DeltekTC/welcome.msv')

Ntime = 5
# ----------------------------------------------------------------------------------------------------------------------
# Introducir dominio
# ----------------------------------------------------------------------------------------------------------------------
WebDriverWait(driver,Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'input#uid'))).send_keys(LoginID)

# ----------------------------------------------------------------------------------------------------------------------
# Introducir contraseña
# ----------------------------------------------------------------------------------------------------------------------
WebDriverWait(driver,Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'input#passField'))).send_keys(Password)

# ----------------------------------------------------------------------------------------------------------------------
# Introducir dominio
# ----------------------------------------------------------------------------------------------------------------------
WebDriverWait(driver,Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'input#dom'))).send_keys(Domain)

# ----------------------------------------------------------------------------------------------------------------------
# Entrar a deltek
# ----------------------------------------------------------------------------------------------------------------------
WebDriverWait(driver,Ntime).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'input#loginButton'))).click()

# ----------------------------------------------------------------------------------------------------------------------
# Borrar los Projects que esten por defecto
# ----------------------------------------------------------------------------------------------------------------------
driver.switch_to.frame(1)
WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "allRowSelector"))).click()
WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "deleteLine"))).click()
time.sleep(0.5)

# ----------------------------------------------------------------------------------------------------------------------
# Diligenciar los Projects - ID
# ----------------------------------------------------------------------------------------------------------------------
for i in range(0,Value["Name"].size):
    # Project ID
    WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "udt" + str(i) + "_1"))).click()
    WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(Value["Project ID"][i])

    # Award ID
    WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "udt" + str(i) + "_3"))).click()
    WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
    WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(str(Value["Award ID"][i]))

    #Activity
    WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "udt" + str(i) +"_4"))).click()
    WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
    WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(str(Value["Activity ID"][i]))
    time.sleep(0.1)

    #driver.switch_to.frame("hrsBodyScroller")
    # driver.switch_to.frame(driver.find_element(By.id("hrsBodyScroller")))
    # WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "hrsBodyScroller"))).click()


for j in range(Datos.columns.size - 1):
    for i in range(0, Value["Name"].size):
        WebDriverWait(driver, Ntime).until(EC.element_to_be_clickable((By.ID, "hrs" + str(i) + "_" + str(j)))).click()
        WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).clear()
        WebDriverWait(driver, Ntime).until(EC.presence_of_element_located((By.ID, "editor"))).send_keys(str(Value.iloc[i,j+4]))
        time.sleep(0.05)

# ----------------------------------------------------------------------------------------------------------------------
# Final processing
# ----------------------------------------------------------------------------------------------------------------------
print("Final - OK")
#¿Qué materias te gustaría descargar?
listC=["ADM-15507", "MAT-12101","MAT-12310","CON-10002",
      "ECO-12102","ADM-12108","EGN-17123","LEN-12702"]

#Paquetes necesarios:
import pandas as pd
import time
from bs4 import BeautifulSoup as bs
from selenium import webdriver

driver = webdriver.Firefox(executable_path="PATH") #Escribe el PATH de geckodriver
driver.get("https://serviciosweb.itam.mx/EDSUP/BWZKSENP.P_Horarios1?s=1809") #Cambia cada periodo
soup=bs(driver.page_source, "html.parser")
listA=[e.text for e in soup("option")] #Las opciones utilizan en #1
listB=[e[0:9] for e in listA] #Extracción de clave de materias

time.sleep(5)

listD=[listB.index(u)+1 for u in listC]
listE=[]

for k in listD:
    boton=driver.find_element_by_xpath("/html/body/div[3]/form/input[2]")
    materiak=driver.find_element_by_xpath('/html/body/div[3]/form/select/option[{}]'.format(k))
    time.sleep(2)
    materiak.click()
    
    time.sleep(7)
    boton.click()
    
    time.sleep(5)
    tablak=pd.read_html(driver.page_source)
    listE.append(tablak)
    driver.back()

driver.quit()

listF=[index for index, value in enumerate(listE)]

writer = pd.ExcelWriter("horarios.xlsx", engine = "xlsxwriter")

n=1
while n!=len(listF)+1:
    listE[n-1][2].to_excel(writer, sheet_name = "Sheet{}".format(n))
    n+=1

writer.save()
writer.close()

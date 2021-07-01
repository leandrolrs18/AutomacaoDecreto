from selenium import webdriver
import time
import xlsxwriter
from datetime import  date
#from cx_Freeze import setup, Executable
import sys

link = 'http://diariooficial.rn.gov.br/dei/dorn3/docview.aspx?id_jor=00000001&data=20210430&id_doc=721589'

web = webdriver.Chrome()
web.get(link)

#print('teste2', web.find_element_by_xpath('/html/body/table/tbody/tr[61]/td[2]/text()').text)
time.sleep(1)
a = web.find_element_by_id('docinf').text
time.sleep(1)
b = web.find_element_by_xpath('/html/body/div/p[5]')
print('a', a)
print('b', b)

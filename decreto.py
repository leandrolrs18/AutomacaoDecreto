from selenium import webdriver
import time
import xlsxwriter
from datetime import  date
#from cx_Freeze import setup, Executable
import sys


def get_all_links(driver):
    links = []
    texto = []
    elements = driver.find_elements_by_tag_name('a')
    for elem in elements:
        href = elem.get_attribute("href")
        texto.append(elem.text)
        links.append(href)
    return links, texto

def start(P_chave, Dincial, Dfinal, paginas):
    links = []
    linkcerto = []
    novolink = []
    info2 = []
    i = 0
    web = webdriver.Chrome()
    web.get('http://diariooficial.rn.gov.br/dei/dorn3/Search.aspx')
    Pchave = web.find_element_by_xpath('//*[@id="input-bs-data"]')
    Pchave.clear()
    Pchave.send_keys(Dincial)
    Pchave = web.find_element_by_xpath('//*[@id="input-bs-data-2"]')
    Pchave.clear()
    Pchave.send_keys(Dfinal)
    Pchave = web.find_element_by_xpath('//*[@id="input-bs-keyword"]')
    Pchave.send_keys(P_chave)
    time.sleep(1)

    Submit = web.find_element_by_xpath('//*[@id="submit-busca-simples"]')
    web.execute_script("arguments[0].click();", Submit)
    #Submit.click()
    #time.sleep(2)

    while True:
        links, texxto = get_all_links(web);   
        #print('texto', texxto)
        for link in links :
            if link is not None:
                if 'docview' in link:
                    linkcerto.append(link)
        t = web.find_element_by_xpath('//*[@id="Form1"]/section[2]/div/div[2]/a[2]')
        t.click() 
        i = i + 1            
        if(i == paginas):
            break

    linkcerto = list(dict.fromkeys(linkcerto))
    #print('a', linkcerto)
    for link in linkcerto:
        web.get(link)
        #time.sleep(2)
        info2.append(web.find_element_by_id('docinf').text)
        d = link.split("data=", 1)[1]
        n = d.split("&", 1)[0]
        l = link.split("doc=", 1)[1]
        novolink.append('http://diariooficial.rn.gov.br/dei/dorn3/documentos/00000001/'+n+'/'+l+'.htm')
    print("qq ", len(info2))
    return novolink, web, info2    

def informacoes(links, web):
    result = []
    for i in links:
        web.get(i)
        #time.sleep(2)
        conteud = web.find_elements_by_class_name("WordSection1")
        for element in conteud:
            result.append(element.text)
    print('qp ', len(result))       
    return result      
    

def filtrar(texto, info2):
    ndata = []
    filtrado = []
    informacoes = []
    info3 = []
    info4 = []
    for i in range(0,len(texto)):
        array = texto[i].split('\n')
        ndata.append(array)
    for j in range(0, len(ndata)):
        if (ndata[j][1].find("DECRETO") != -1) :
            filtrado.append(ndata[j])
            info3.append(info2[j])
    #print('filtrados', filtrado)
    print('t', filtrado)
    for j in range(0, len(filtrado)):
        informacoes.append((filtrado[j][1].split(",", 1)[0]))
        informacoes.append(filtrado[j][2])

    print ('info', informacoes)
    for j in range(0, len(info3)):
       a = info3[j].split("em:", 1)[1]
       b = a.split("Edição", 1)[0]
       c = info3[j].split("Diária:", 1)[1]
       info4.append(b)
       info4.append(c)


    return filtrado, informacoes, info4

def gerarExcel (texto, texto2):
    today = date.today()
    today = str(today) +'.xlsx'
    #print(today)
    with xlsxwriter.Workbook(today) as workbook:
        worksheet = workbook.add_worksheet()
        cont = 0
        cont1 = 0
        for row_num, data in enumerate(texto):
            print(row_num)#, str(data))
            for i in range(0, 4):
                if(i<2) :
                    worksheet.write_string(row_num, i , str(texto[cont]))
                    cont = cont + 1
                else:
                    worksheet.write_string(row_num, i , str(texto2[cont1]))    
                    cont1 = cont1 + 1
            if row_num == (len(texto)/2-1):
                break
    
if __name__ == '__main__':
    
    links = []
    texto = []
    info2 = []
    info3 = []
    links, web, info2 = start("Decreto nº fátima bezerra", "01/01/2015", "02/07/2021", 340)  # parametros: palavra de pesquisa e numero de pag pesquisadas 
    texto = informacoes(links, web)  
    filtrados, info, info3 = filtrar(texto, info2)
    print('info3', info3)
    #separar info3
    gerarExcel(info, info3)
    



# o info é o dobro de info 3 porque tem decreto e emenda, info3 tem q separar em nº e data
#otimizar código. talvez criar classe quando conseguir generalizar as funções
#falta criar um executável - qt pode ser uma opção para aplicação, mas posso criar um agendamento
#criar titulos de colunas no excel

#erro comum: quando volta o elemento, tem que ter o ".text" no findby...
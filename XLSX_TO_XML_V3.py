import xml.etree.ElementTree as ET
from xml.dom import minidom
from PySimpleGUI.PySimpleGUI import T
import numpy as np
import pandas as pd
from datetime import datetime
import PySimpleGUI as sg   
import os
import paramiko
import socket
from socket import AF_INET, SOCK_DGRAM




def sendXML(server, username, password, localpath, remotepath):
    print(server, username, password, localpath, remotepath)
    ssh = paramiko.SSHClient() 
    ssh.load_host_keys(os.path.expanduser(os.path.join("~", ".ssh", "known_hosts")))
    ssh.connect(server, username=username, password=password, timeout=10)
    sftp = ssh.open_sftp()
    sftp.put(localpath, remotepath)
    sftp.close()
    ssh.close()

def readXLSX(filename):
    xl = pd.read_excel(filename)
    cols = range(1,31)
    for elem in cols:elem = str(elem)
    xl = xl.set_axis(cols, axis=1, inplace=False)
    data = {}
    shop = {"name":xl[1][0],
            "company":xl[2][0],
            "url":xl[3][0]}
    data["shop"] = shop

    currencies = {}
    if len(xl[4].dropna()) == len(xl[5].dropna()):
        for i in range(1,len(xl[4].dropna())):
            currencies[str(xl[4].dropna()[i])] = str(xl[5].dropna()[i])
    else:
        print("bad")
    data["currencies"] = currencies

    categories = []
    if len(xl[6].dropna()) == len(xl[8].dropna()):
        for i in range(1,len(xl[6].dropna())):
            if xl[7][i] is np.nan:
                categories.append({"id": str(xl[6].dropna()[i]), "text": xl[8][i] })
            else:
                categories.append({"id": str(xl[6].dropna()[i]), "parentId": str(xl[7][i]), "text": xl[8][i] })
    else:
        print("bad")
    data["categories"] = categories

    shipment_options = []
    if len(xl[9].dropna()) == len(xl[10].dropna()):  
        for i in range(1,len(xl[9].dropna())):
            if xl[11][i] is np.nan:
                shipment_options.append({"days":str(xl[10].dropna()[i]),
                                         "order-before": str(xl[9].dropna()[i])})
            else:
                shipment_options.append({"days": str(xl[10].dropna()[i]),
                                         "id": str(xl[11][i]),
                                         "order-before":str(xl[9].dropna()[i])})
    else:
        print("bad")

    data["shipment_options"] = shipment_options

    offers = []

    offers_position = []
    if len(xl[12].dropna()) >= 1:
        for i in range(1,  len(xl[12])):
            if  xl[12][i] is not np.nan:
                offers_position.append(i)

    if len(offers_position) == 1:
        offer_begin = 1
        offer_end = len(xl[12])
                
        offer = {"attributes":{},"params_without_attrs":{}, "shipment-options":[], "outlets":[], "params":{}}

        offer["attributes"]["id"] = str(xl[12][offer_begin])
        if xl[13][offer_begin] is not np.nan:
            if xl[13][offer_begin] == "Да":
                offer["attributes"]["available"] = "true"
            else:
                offer["attributes"]["available"] = "false"
        
        if xl[14][offer_begin] is not np.nan:
            offer["params_without_attrs"]["url"] = str(xl[14][offer_begin])
        
        offer["params_without_attrs"]["name"] = str(xl[15][offer_begin])
        offer["params_without_attrs"]["price"] = str(xl[16][offer_begin])
        offer["params_without_attrs"]["categoryId"] = str(xl[17][offer_begin])

        if xl[18][offer_begin] is not np.nan:
            offer["params_without_attrs"]["picture"] = str(xl[18][offer_begin])

        if xl[19][offer_begin] is not np.nan:
            offer["params_without_attrs"]["vat"] = str(xl[19][offer_begin])

        if xl[22][offer_begin] is not np.nan:
            offer["params_without_attrs"]["vendor"] = str(xl[22][offer_begin])

        if xl[23][offer_begin] is not np.nan:
            offer["params_without_attrs"]["vendorCode"] = str(xl[23][offer_begin])

        if xl[24][offer_begin] is not np.nan:
            offer["params_without_attrs"]["model"] = str(xl[24][offer_begin])

        if xl[25][offer_begin] is not np.nan:
            offer["params_without_attrs"]["description"] = str(xl[25][offer_begin])

        if xl[26][offer_begin] is not np.nan:
            offer["params_without_attrs"]["barcode"] = str(xl[26][offer_begin])
        
        for j in range(offer_begin, offer_end):
            if xl[20][j] is not np.nan:
                offer["shipment-options"].append({"days": str(xl[20][j]),  "order-before": str(xl[21][j])})
            
        for j in range(offer_begin, offer_end):
                if xl[27][j] is not np.nan:
                    offer["outlets"].append({"id":  str(xl[27][j]),  "instock": str(xl[28][j])})

        for j in range(offer_begin, offer_end):
            if xl[29][j] is not np.nan:
                offer["params"][str(xl[29][j])] = str(xl[30][j])

        offers.append(offer)
    else:
        for i in range(0, len(offers_position) - 1):
            #offers without last
            offer = {"attributes":{},"params_without_attrs":{}, "shipment-options":[], "outlets":[], "params":{}}
            offer_begin = offers_position[i]
            offer_end = offers_position[i + 1]
            end_of_last_offer = len(xl[12])

            offer["attributes"]["id"] = str(xl[12][offer_begin])
            if xl[13][offer_begin] is not np.nan:
                if xl[13][offer_begin] == "Да":
                    offer["attributes"]["available"] = "true"
                else:
                    offer["attributes"]["available"] = "false"
            
            if xl[14][offer_begin] is not np.nan:
                offer["params_without_attrs"]["url"] = str(xl[14][offer_begin])
            
            offer["params_without_attrs"]["name"] = str(xl[15][offer_begin])
            offer["params_without_attrs"]["price"] = str(xl[16][offer_begin])
            offer["params_without_attrs"]["categoryId"] = str(xl[17][offer_begin])

            if xl[18][offer_begin] is not np.nan:
                offer["params_without_attrs"]["picture"] = str(xl[18][offer_begin])

            if xl[19][offer_begin] is not np.nan:
                offer["params_without_attrs"]["vat"] = str(xl[19][offer_begin])

            if xl[22][offer_begin] is not np.nan:
                offer["params_without_attrs"]["vendor"] = str(xl[22][offer_begin])

            if xl[23][offer_begin] is not np.nan:
                offer["params_without_attrs"]["vendorCode"] = str(xl[23][offer_begin])

            if xl[24][offer_begin] is not np.nan:
                offer["params_without_attrs"]["model"] = str(xl[24][offer_begin])

            if xl[25][offer_begin] is not np.nan:
                offer["params_without_attrs"]["description"] = str(xl[25][offer_begin])

            if xl[26][offer_begin] is not np.nan:
                offer["params_without_attrs"]["barcode"] = str(xl[26][offer_begin])
            
            for j in range(offer_begin, offer_end):
                if xl[20][j] is not np.nan:
                    offer["shipment-options"].append({"days":  str(xl[20][j]),  "order-before": str(xl[21][j])})
            
            for j in range(offer_begin, offer_end):
                if xl[27][j] is not np.nan:
                    offer["outlets"].append({"id":  str(xl[27][j]),  "instock": str(xl[28][j])})

            for j in range(offer_begin, offer_end):
                if xl[29][j] is not np.nan:
                    offer["params"][xl[29][j]] = str(xl[30][j])

            offers.append(offer)

        #last offer in many offers
        offer = {"attributes":{},"params_without_attrs":{}, "shipment-options":[], "outlets":[], "params":{}}
        offer_begin = offers_position[-1]
        end_of_last_offer = len(xl[12])
        offer["attributes"]["id"] = str(xl[12][offer_begin])
        if xl[13][offer_begin] is not np.nan:
            if xl[13][offer_begin] == "Да":
                offer["attributes"]["available"] = "true"
            else:
                offer["attributes"]["available"] = "false"
            
        if xl[14][offer_begin] is not np.nan:
            offer["params_without_attrs"]["url"] = str(xl[14][offer_begin])
            
        offer["params_without_attrs"]["name"] = str(xl[15][offer_begin])
        offer["params_without_attrs"]["price"] = str(xl[16][offer_begin])
        offer["params_without_attrs"]["categoryId"] = str(xl[17][offer_begin])

        if xl[18][offer_begin] is not np.nan:
            offer["params_without_attrs"]["picture"] = str(xl[18][offer_begin])

        if xl[19][offer_begin] is not np.nan:
            offer["params_without_attrs"]["vat"] = str(xl[19][offer_begin])

        if xl[22][offer_begin] is not np.nan:
            offer["params_without_attrs"]["vendor"] = str(xl[22][offer_begin])

        if xl[23][offer_begin] is not np.nan:
            offer["params_without_attrs"]["vendorCode"] = str(xl[23][offer_begin])

        if xl[24][offer_begin] is not np.nan:
            offer["params_without_attrs"]["model"] = str(xl[24][offer_begin])

        if xl[25][offer_begin] is not np.nan:
            offer["params_without_attrs"]["description"] = str(xl[25][offer_begin])

        if xl[26][offer_begin] is not np.nan:
            offer["params_without_attrs"]["barcode"] = str(xl[26][offer_begin])
            
        for j in range(offer_begin, end_of_last_offer):
            if xl[20][j] is not np.nan:
                offer["shipment-options"].append({"days":  str(xl[20][j]),  "order-before": str(xl[21][j])})
            
        for j in range(offer_begin, end_of_last_offer):
            if xl[27][j] is not np.nan:
                offer["outlets"].append({"id":  str(xl[27][j]),  "instock": str(xl[28][j])})

        for j in range(offer_begin, end_of_last_offer):
            if xl[29][j] is not np.nan:
                offer["params"][xl[29][j]] = str(xl[30][j])
        offers.append(offer)

    data["offers"] = offers

    print(data)

    return(data)


def createXML(filename, data):
    now = datetime.now()
    current_time = now.strftime("%Y-%m-%d %H:%M")
    yml_catalog = ET.Element("yml_catalog", attrib={"date": current_time})
    shop = ET.SubElement(yml_catalog,"shop")
    shop_attr = data["shop"]
    name =  ET.SubElement(shop, "name")
    name.text = shop_attr["name"]
    company =  ET.SubElement(shop, "company")
    company.text = shop_attr["company"]
    url =  ET.SubElement(shop, "url")
    url.text = shop_attr["url"]


    if data["currencies"] is not {}:
        currencies =  ET.SubElement(shop, "currencies")
        currencies_attr = data["currencies"]
    #currencies_array = {'RUR': "1", 'USD': "80"}
        for elem in currencies_attr:
            currency = ET.SubElement(currencies, "currency")
            currency.set(elem, currencies_attr[elem])


    categories =  ET.SubElement(shop, "categories")
    categories_attr = data["categories"]
    #categories_array = [{"id":"1278","text":"Электроника"},{"id":"3761", "parentId":"1278", "text": "Телевизоры"},
    #{"id":"1553", "parentId":"3761", "text": "Медиаплееры"}, {"id": "3798","text":"Бытовая техника"},
    #{"id":"1293", "parentId":"3798", "text": "Холодильники"}]
    for elem in categories_attr:
        category = ET.SubElement(categories, "category")
        category.set("id", elem["id"])
        if "parentId" in elem.keys():
            category.set("parentId", elem["parentId"])
        category.text = elem["text"]



    shipment_options =  ET.SubElement(shop, "shipment-options")
    shipment_options_attr = data["shipment_options"]
    #shipment_options_array = [{"days":"1", "order-before": 15}]

    for elem in shipment_options_attr:
        option = ET.SubElement(shipment_options, "option")
        option.set("days", elem["days"])
        option.set("order-before", elem["order-before"])
    
    offers = ET.SubElement(shop, "offers")
 
    offers_attr = data["offers"]
    #offers_array = [{"attributes":{"id":"158", "available":"true"}, 
    #"params_without_attrs":{"url":"www.abc.ru", "name":"Холодильник Индезит СБ 185", "price":"18500", "categoryId":"1293", "picture":"www.abc.ru/img.jpg", "vat":"2",
    #    "vendor":"Indesit", "vendorCode":"12345678", "model":"Indesit SB 185", "description":"Холодильник Indesit SB 185", "barcode":"7564756475648"},
    #"shipment-options":[{"days":"1", "order-before":15},{"days":"3", "order-before":7}],
    #"outlets":[{"id":"1", "instock":"50"},{"id":"2", "instock":22}],
    #"params":{"Weight":"120", "Width":"70", "Length": "250"}},]

    for elem in offers_attr:
        offer = ET.SubElement(offers, "offer")
        offer.set("id", elem["attributes"]["id"]) 
        offer.set("available", elem["attributes"]["available"])

        for param in elem["params_without_attrs"]:
            unit = ET.SubElement(offer, param)
            unit.text = elem["params_without_attrs"][param]

        offer_shipment_options = ET.SubElement(offer, "shipment-options")
        for option in elem["shipment-options"]:
            opt = ET.SubElement(offer_shipment_options,"option")
            for par in option:
                opt.set(par, option[par])

        offer_outlets = ET.SubElement(offer, "outlets")
        for outlet in elem["outlets"]:
            out = ET.SubElement(offer_outlets, "outlet")
            for par in outlet:
                out.set(par, outlet[par])

        for param in elem["params"]:
            par = ET.SubElement(offer, "param")
            par.set("name", param)
            par.text = elem["params"][param]

    tree = ET.tostring(yml_catalog, encoding='utf-8',method='xml', xml_declaration=True)


    dom = minidom.parseString(tree)
    final = dom.toprettyxml(indent='\t')
    
    file = open(filename, "w", encoding="utf-8")
    file.write(final)
    file.close()
    file = open(filename, encoding="utf-8")
    lines = file.readlines()
    lines[0] = '<?xml version="1.0" encoding="UTF-8"?>\n'
    with open(filename, "w", encoding="utf-8") as file:
        file.writelines(lines)
    file.close()
if __name__ == "__main__":
    sg.theme('Light Blue 2')    

    layoutConvert = [
            [sg.Text('Выберите файл Exel', font=("Arial", 11))],
            [sg.Input(key='-USER FOLDER-'), sg.FileBrowse(target='-USER FOLDER-', key='BROWSE', font=("Arial", 11))],
            [sg.Button('Конвертировать', font=("Arial", 11))]
            ]

    layoutSend = [
            [sg.Text('IP адрес', font=("Arial", 11))],
            [sg.Input(key='IP')],
            [sg.Text('Пользователь', font=("Arial", 11))],
            [sg.Input(key='USER')],
            [sg.Text('Пароль', font=("Arial", 11))],
            [sg.Input(key='PASS')],
            [sg.Text('Выберите локальный файл XML', font=("Arial", 11))],
            [sg.Input(key='XML_PATH'), sg.FileBrowse(target='XML_PATH', key='BROWSE_XML', font=("Arial", 11))],
            [sg.Text('Файл удалённого сервера', font=("Arial", 11))],
            [sg.Input(key='REMOTE_DIR')],
            [sg.Button('Отправить', font=("Arial", 11))]
            ]
            
    layout = [
        [sg.TabGroup([[sg.Tab('Конвертировать в XML', layoutConvert, tooltip='tip1'), sg.Tab('Отправить на сервер', layoutSend, tooltip='tip2')]])]
        ]

    window = sg.Window('Excel to XML', layout)
    while True:
        try:
            event, values = window.read()


            if event in (sg.WIN_CLOSED, 'Exit'):
                break


            if event == "Конвертировать":
                folder = sg.popup_get_folder("Выберите дирректорию для сохранения файла.", font=("Arial", 11))
                filename = values['-USER FOLDER-']
                createXML(folder+"/СберМегаМаркет.xml", readXLSX(filename))
                sg.popup("Выполнено! Файл xml находиться в папке " + folder, font=("Arial", 11))

            if event == "Отправить":
                if values["IP"] == "":
                    sg.popup_error("Введите IP!")
                elif values["USER"] == "":
                    sg.popup_error("Введите имя пользователя!")
                elif values["PASS"] == "":
                    sg.popup_error("Введите пароль!")
                elif values["XML_PATH"] == "":
                    sg.popup_error("Выберите файл XML!")
                elif values["REMOTE_DIR"] == "":
                    sg.popup_error("Введите полный путь к файлу сервера!")
                else:
                    sendXML(values["IP"], values["USER"], values["PASS"], values["XML_PATH"], values["REMOTE_DIR"])
                    sg.popup("Отправлено!")

        except ValueError:
            sg.popup_error("Выбран файл с неподходящим расширением!", font=("Arial", 11))


        except FileNotFoundError:
            sg.popup_error("Пожалуйста, выберите файл!", font=("Arial", 11))

        except socket.timeout:
            sg.popup_error("Превышено время ожидания соединения. Проверьте доступность сервера!")

        except paramiko.ssh_exception.AuthenticationException:
            sg.popup_error("Ошибка аутентификации! Проверьте имя пользователя и пароль!")
    window.close()
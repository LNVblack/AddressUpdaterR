import time
import os
import win32com.client
import random
import requests
import json

class Address_Updater:
    '''
        Принимает на вход:
            sheet - Имя листа
            str_row - номер строки с которой будет начинаться считывание таблицы
            int_column - номер столбца который будет считываться
            last_row - номер строки на которой работа скрипта будет заканчиваться
            column_save_adr - столбец для сохранения обновленных адресов
            post_str - столбец для сохранения почтовог индекса
    '''
    def __init__(self,workbook, sheet = 'Лист1', str_row = 1, int_column = 1, last_row = 1, column_save_adr = 1, post_str = 1):
        self.workbook = workbook
        self.sheet = sheet
        self.str_row = str_row
        self.int_column = int_column
        self.last_row = last_row
        self.column_save_adr = column_save_adr
        self.post_str = post_str
        self.error_list = [
            'Ошибка, адрес заполнен не правильно',
            'Hазвание страны не получено. :c',
            'Ошибка, название округа не получено',
            'Ошибка, название области не получено',
            'Ошибка, названиe города не получено',
            'Hазвание страны не получено. :c, Ошибка, название области не получено, Ошибка, названиe города не получено,  ,  ']
        self.GY = GeoYandex("fc9fd2bb-871a-4af7-8cd2-6ecfe936e669")

    def Address_Update(self):
        '''
            Считывает данные из ячейки в таблице Excel и отправляет запрос в Yandex API (Geocoder)
            После получения ответа на запрос в формате JSON, вызывет функцию "Assembly_Update" и "Post_code" 
        '''
        for i in range(self.str_row, self.last_row+1):
            if i not in ['Адрес','NULL','None']:
                self.GY.SendGetQuery(str(self.sheet.Cells(self.str_row,self.int_column).value)) #Отправка запроса в YANDEX
                self.sheet.Cells(self.str_row,self.column_save_adr).value = self.Assembly_Address()
                self.Post_Code(self.post_str)
                self.str_row += 1
    
    def Assembly_Address(self):
        '''
            Распарсивает JSON файл и получает данные в формате: Россия Смоленская область Смоленск улица Матросова 27
        '''
        new_address = (( (self.GY.ExtractCountry()) + ', ' + (self.GY.ExtractRegion()) + ', ' + (self.GY.ExtractLocality()) + ', ' + (self.GY.ExtractStreet())+ ', ' + (self.GY.ExtractHouse()) ))
        if new_address in self.error_list:
            return self.Address_Fix()
        else:
            return new_address

    def Address_Fix(self):
        '''
            При получении в пустого JSON срезает последнее слово и отправляет новый запрос
            Делает это до тех пор, пока не будет получен удовлетворительный JSON или пока строка запроса не станет пустой  
        '''
        print('Обработка исключения в строке '+str(self.str_row))
        er_adr = self.sheet.Cells(self.str_row,self.int_column).value
        slicer = er_adr.count(' ')
        parts = er_adr.rsplit(' ', slicer)
        cnt = 1
        for i in parts:
            self.GY.SendGetQuery(str(parts[0:-cnt]))
            fix_adr = (( (self.GY.ExtractCountry()) + ', ' + (self.GY.ExtractRegion()) + ', ' + (self.GY.ExtractLocality()) + ', ' + (self.GY.ExtractStreet())+ ', ' + (self.GY.ExtractHouse()) ))
            cnt += 1
            if fix_adr not in self.error_list:
                self.sheet.Cells(self.str_row,self.column_save_adr).value = str(fix_adr)
                break
            else:
                cnt += 1
        return str(fix_adr)

    def Post_Code(self, post_str):
        '''
            Распарсивает JSON файл и получает почтовый индекс:
        '''
        postcode = self.GY.ExtractPostalCode()
        self.sheet.Cells(self.str_row,self.post_str).value = str(postcode)

    def Save_and_Quit(self):
        '''
            Сохраняет файл и закрывает его
        '''
        print('Обработка закончена')
        self.workbook.Save()

class GeoYandex:
    '''
        GeoYandex - класс, содержащий методы и переменные для упрощённого взаимодействия с Yandex Geocoder API
        На вход нужно дать параметр Ключ API

        -YaAPI = "01234567-89ab-cdef-0123-456789abcdef"

        Стандартные параметры:
        -baseurl = "http://geocode-maps.yandex.ru/1.x" - базовая ссылка на Яндекс. Менять...смысла нет
        -format = 'json' - формат получения данных. Все методы работают по нему.
        -lang = 'ru'     - язык полученных данных
        -results = '1'
        '''
    def __init__(self, YaAPI):
        self.baseurl = "http://geocode-maps.yandex.ru/1.x"
        self.YaAPI = YaAPI
        self.format = 'json'
        self.lang = 'ru'
        self.results = '1'
        self.PARAMS = {'apikey' : self.YaAPI, 
                       'format' : self.format,
                       'lang'   : self.lang,
                       'results': self.results }
    
    def SendGetQuery(self, Address):
        '''
        Метод для отправления запроса. Следует использовать перед попыткой получения(парсинга) данных
        Список методов получения:
        -ExtractAddress()  - Получение Адреса по версии Яндекса
        -ExtractPostalCode() - Получение Почт. Индекса
        -ExtractCoordinates() - Получение Координат через пробел
        '''
        time.sleep(0.5)
        self.PARAMS['geocode'] = Address
        self.result = requests.get(url = self.baseurl, params= self.PARAMS)
        self.data = self.result.json()
        #print(self.data)

    def ExtractAddress(self):
        '''
        Получение Адреса по версии Яндекса
        -Пример: Россия, Москва, улица Льва Толстого, 16
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['text']
        except IndexError:
            return 'Ошибка, адрес заполнен не правильно' 

    def ExtractCountry(self):
        '''
            Получение названия страны
            -Пример: Россия
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][0]['name']
        except IndexError:
            return 'Hазвание страны не получено. :c'

    def ExtractProvince(self):
        '''
            Получение названия округа
            -Пример: Приволжский федеральный округ
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][1]['name']
        except IndexError:
            return 'Ошибка, название округа не получено'
    
    def ExtractRegion(self):
        '''
            Получение названия области
            -Пример: Самарская область
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][2]['name']
        except IndexError:
            return 'Ошибка, название области не получено'
    
    def ExtractArea(self):
        '''
            Получение названия областного округа
            -Пример: городской округ Самара
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][3]['name']
        except IndexError:
            return 'Ошибка, название областного округа не получено'

    def ExtractLocality(self):
        '''
            Получение названия города
            -Пример: Самара
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][4]['name']
        except IndexError:
            return 'Ошибка, названиe города не получено'

    def ExtractStreet(self):
        '''
            Получение названия улицы
            -Пример: улица Георгия Димитрова
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][-2]['name']
        except IndexError:
            return ' '

    def ExtractHouse(self):
        '''
            Получение цифры дома
            -Пример: 44
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components'][-1]['name']
        except IndexError:
            return ' '

    def ExtractPostalCode(self):
        '''
        Получение Почт. Индекса
        -Пример: 119021
        '''
        try:
            return self.data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['postal_code']
        except LookupError:
            return 'Ошибка, почт. индекс не получен'

class Logo:
    def __init__(self):
        self.rand = random.randint(0, 2)

    def Logo_rand(self):
        if self.rand == 0:
            self.Logo_0()
        elif self.rand == 1:
            self.Logo_1()
        else:
            self.Logo_2()
        
    def Logo_0(self):
        print('\n\n\
        ╔══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╗\n\
        ╚══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╝\n\
        ──────────────────────────────────────────────────────────────────────\n\n\
        ╔══╗╔══╗╔═══╦═══╦══╦══╦═══╦══╦╗──╔╦═══╦╗╔╦═══╦════╦══╦═══╦══╦══╦══╦══╗\n\
        ║╔╗║║╔╗║║╔═╗║╔══╣╔═╣╔╗╠══╗║╔╗║║──║║╔══╣║║║╔═╗╠═╗╔═╣╔╗║╔═╗║╔╗║╔╗║╔╗║╔╗║\n\
        ║╚╝║║║║║║╚═╝║╚══╣║─║║║║╔═╝║╚╝║╚╗╔╝║╚══╣╚╝║╚═╝║─║║─║║║║╚═╝║╚╝║║║║║║║║║║\n\
        ║╔╗║║║║║║╔══╣╔══╣║─║║║║╚═╗║╔╗║╔╗╔╗║╔══╣╔╗╠╗╔╗║─║║─║║║║╔══╩═╗║║║║║║║║║║\n\
        ║║║╠╝╚╝╚╣║──║╚══╣╚═╣╚╝╠══╝║║║║║╚╝║║╚══╣║║║║║║║─║║─║╚╝║║──╔═╝║╚╝║╚╝║╚╝║\n\
        ╚╝╚╩════╩╝──╚═══╩══╩══╩═══╩╝╚╩╝──╚╩═══╩╝╚╝╚╝╚╝─╚╝─╚══╩╝──╚══╩══╩══╩══╝\n\n\
        ──────────────────────────────────────────────────────────────────────\n\
        ╔══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╦══╗\n\
        ╚══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╩══╝\n\n')
    def Logo_1(self):
        print('\n\n\
        ████████████████████████\n\
        █─███─█────█────█─██─█─█\n\
        █─███─█─██─█─██─█─█─██─█\n\
        █─█─█─█─██─█────█──███─█\n\
        █─────█─██─█─█─██─█─████\n\
        ██─█─██────█─█─██─██─█─█\n\
        ████████████████████████\n\n')
    def Logo_2(self):
        print('\n\n\
        $$$$$__$$__$$_$$$$$$_$$__$$__$$$$__$$__$$\n\
        $$__$$__$$$$____$$___$$__$$_$$__$$_$$$_$$\n\
        $$$$$____$$_____$$___$$$$$$_$$__$$_$$_$$$\n\
        $$_______$$_____$$___$$__$$_$$__$$_$$__$$\n\
        $$_______$$_____$$___$$__$$__$$$$__$$__$$\n\n')


if __name__ == "__main__":
    Logo = Logo()
    Logo.Logo_rand()

    Excel = win32com.client.Dispatch("Excel.Application")                                               #Создание COM объекта

    while True:
        filepath = str(input('укажите путь к файлу, или введите "0" для закрытия программы \n >>> '))   #Ввод пути и имени файла
        filepath += '.xlsx'                                                                             #Добавление к строке типа файлв (xlsx)
        if os.path.exists(filepath):                                                                    #Если файл по указанному пути сузествует то...
            Excel.Visible = False                                                                       #Открытие файла Excel в режиме отображения
            workbook = Excel.WorkBooks.Open(u''+filepath)                                               #Открытие рабочей книги
            print('\nАктивная книга: '+str(workbook.Name))                                              #Вывод полного имени рабочей книги

            time.sleep(0.5)                                                                             #Задержка 0,5 сек

            menu = str(input('\
                1: Работать с активной книгой\n\
                2: Выбрать другую книгу\n\
                3: Закрыть программу\n\
                >>> '))                                                                                 #Меню действий с книгой

            if str(menu) == '1':                                                                        #Если выбран первый пункт меню...
                List_Sheets = []                                                                        #Пустой список
                for i in workbook.Sheets:                                                               #Цикл прохода по рабочей книге с целью поиска всех листов
                    List_Sheets.append(i.Name)                                                          #Добавление в список Имени листа
                
                cnt = 0                                                                                 #Счетчик
                for i in List_Sheets:                                                                   #Цикл прохода по списку наименований листов
                    print (str(cnt)+') '+str(i))                                                        #Вывод наименования листа
                    cnt+=1                                                                              #Счетчик + 1
                
                Sheet_Number = int(input('Введите номер листа: \n >>> '))                               #Ввод номера листа
    
                if Sheet_Number in range(0, len(List_Sheets)):                                          #Если Sheet_number в диапазоне List_Sheets 
                    sheet = workbook.Worksheets(u''+str(List_Sheets[Sheet_Number]))

                    str_row = int(input('Введите цифру строки для начала считывания \n >>> '))          #Ввод цифры строки с которой будут считываться данные
                    int_column = int(input('Введите цифру столбца для начала считывания\n >>> '))       #Ввод цифры столбца в диапазоне которго будет проходить считывание
                    last_row = int(input('Введите цифру строки для окончания считывания\n >>> '))       #Ввод цифры строки на которой программа закончит считывание
                    column_save_adr = int(input('Введите цифру столбца для записи адреса\n >>> '))      #Ввод цифры столбца для записи нового адреса 
                    post_str = int(input('Введите цифру столбца для записи почтового индеска\n >>> '))  #Ввод цифры столбца для записи почт. индекса

                    Address_Up_Class = Address_Updater(workbook, sheet, str_row, int_column, last_row, column_save_adr, post_str)   #Создание объекта принадлежащего к классу Address_Updater
                    Address_Up_Class.Address_Update()
                    Address_Up_Class.Save_and_Quit()

                else: 
                    print('Выбранного листа не существует...')
                    time.sleep(1)

            elif str(menu) == '2':
                workbook = Excel.WorkBooks.Open(u''+str(input('Путь к книге \n >>> ')))

            elif str(menu) == '3':
                Excel.Quit
                break

            else:
                continue
        
        elif filepath == '0.xlsx':
            print('Выход...')
            time.sleep(0.5)
            break
        else:
            print('Файл не найден...')
            time.sleep(1)
            continue
#Подключаем библиотеки #import excel as excel #Наш файл для работы с екселем
import keyboard
from selenium import webdriver #Библиотека для работы с Браузером
from selenium.webdriver.common.by import By # облегчаем работа с драйвером
import pandas as pd #Библиотека для работы с таблицами
def get_table_excel(path):
    """Функция получения таблицы с ексель с первого листа"""
    table_excel = pd.read_excel(path)
    return table_excel

try:
    settings_table = get_table_excel(".//Setting.xlsx") # получаем таблицу с настройками и путями
    table = get_table_excel(settings_table["Value"][0]) # получаем таблицу с ексель файла
    browser = webdriver.Chrome() # Запускаем браузер
    print("Запускаем браузер")
    browser.get(settings_table["Value"][1]) # Открываем сайт
    print("Открываем сайт:"+settings_table["Value"][1])
    assert "Rpa Challenge" in browser.title # смотрим что название сайта правильное
    browser.find_element(By.XPATH,settings_table["Value"][6]).click() # запускаем тест
    print("Начало теста")
    i = 0 # счетчик
    for val in table.values: # цикл для каждой строки в таблице
        print( "форма "+str(i+1) +":-----------------------------------------------------------------")
        form_element = browser.find_elements(By.TAG_NAME,settings_table["Value"][2]) # Получаем форму
        submit_element = browser.find_element(By.XPATH,settings_table["Value"][3]) # Получаем кнопку отправки
        # Перебираем форму на инпуты
        for element in form_element:
            name_input = element.find_element(By.TAG_NAME,settings_table["Value"][4]).text.strip() # получаем имя инпута``
            name_columns = [x for x in table.columns.to_list() if x.__contains__(name_input)] # смотрим,что в нашей таблице имеется такое имя.ТК в ексель файле могут быть артефакты(лишний пробел и тд).
            input_element = element.find_element(By.TAG_NAME,settings_table["Value"][5]) # находим импукт
            input_element.send_keys(str(table[name_columns[0]][i])) # заполняем инпукт
            print(name_input + ": "+str(table[name_columns[0]][i]))
        i = i + 1 # счетчик
        submit_element.click() # отправить форму
        print("отправили форму!")
except Exception as ex:
        print(ex)
finally:
    browser.save_screenshot('screenie.png')
    print("Конец теста")
    input("Press keyboard to end...")
    keyboard.read_key() 
    browser.close()
    exit()

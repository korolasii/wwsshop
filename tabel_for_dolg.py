import openpyxl
import base_json_test
import base_dolg
import PySimpleGUI as sg
import time


def header_table_for_dolg(way='Main_new.xlsm'):
    book = openpyxl.open(str(way), read_only=True)  # Открытие файла
    sheet = book.worksheets[book.sheetnames.index('Лист5')]  # Запуск листа с названием - получение
    return dolg(base_dolg.base_dolg_reload(way), sheet), dolg_diferent(), mylist_def(), message_for_header_table_dolg(sheet)

#Собираю данные о долгах
def dolg(dolg_dickt, sheet):
    global dolg_list
    dolg_list={}
    for key, value in dolg_dickt.items():
        if value <= int(sheet[1][0].value):
            dolg_list[key]=value
    return dolg_list

#Создаю список людей
def dolg_diferent():
    global gmail_list
    gmail_baze = base_json_test.base_buer()
    gmail_list = {}
    for i in dolg_list.keys():
        if i.strip() not in gmail_list.keys() and i.strip() in gmail_baze.keys():
            gmail_list[i.strip()] = gmail_baze[i.strip()]
    return gmail_list


#Создаю шапку письма
def message_for_header_table_dolg(file_exel):
    global message
    sg.one_line_progress_meter('Рассылка', mylist[1], len(gmail_list.keys()) * 2 + 2, 'Собераем даные о письме')
    time.sleep(1)
    message = {}
    key = 'a'
    for i in range(2, file_exel.max_row+1):
        if file_exel[i][0].value != None:
            key += 'a'
            message[key] = file_exel[i][0].value
    return message

#Создал прогрес бар
def mylist_def():
    global mylist
    mylist = [1,2]
    for i in range(len(gmail_list.keys())*2):
        mylist.append(mylist[len(mylist)-1]+1)
    sg.one_line_progress_meter('Рассылка', mylist[0], len(gmail_list.keys()) * 2 + 3,'Собираем даные о долгах')
    time.sleep(1)
    return mylist


if __name__ == '__main__':
    print(header_table_for_dolg())
import json
import openpyxl

def base_byer_reload(way='Main_new.xlsm'):
    data={}
    book = openpyxl.open(str(way), read_only=True)  # Открытие файла
    gmail_baze = book.worksheets[book.sheetnames.index('База клиентов')]
    for i in range(2,gmail_baze.max_row+1):
        data[gmail_baze[i][0].value.strip()]=gmail_baze[i][1].value.strip()
    open_file_json_w(data)
    return data

def base_buer():
    with open('base.json', 'r') as file:
        data=json.load(file)
    return data


def open_file_json_w(data):
    with open('base.json', 'w') as file:
       json.dump(data, file, indent=4)
    return data


if __name__ == '__main__':
#    print(base_byer_reload())
    print(base_buer().keys())


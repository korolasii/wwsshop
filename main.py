from send_emeil import main_send_email
from table import header_table
from jinja2 import Environment, FileSystemLoader
import PySimpleGUI as sg
import time



def html_header_body(message_header, message_header_tabel, message_body_table, buers, dolg):
    file_loader = FileSystemLoader('')
    env = Environment(loader=file_loader)

    tm = env.get_template('web/index_1.html')
    return tm.render(message=message_header, message_header_tabel=message_header_tabel,
                     message_body_table=message_body_table, buers=buers, dolg=dolg)




def main(sendler, password, way):
    subject, message_header, list_emeil, progres_bar, message_header_tabel, message_body_table, dolg = header_table(way)
    progres_bar_chet = 2
    list_emeil_keys = list_emeil.keys()
    error = []
    for i in list_emeil_keys:
        if i in dolg:
            continue
        else:
            error.append(i)
            break
    if len(error) != 0:
        sg.one_line_progress_meter('Рассылка', len(list_emeil.keys()) * 2 + 3, len(list_emeil.keys()) * 2 + 3,'Ошибка')
        layout = [[sg.Text("База клиентво не совпадает с долгами")],
                  [sg.Text(f'Проверти ник клиентa {error[0]} и наличие иго в долгах')],
                  [sg.Text('Для выхода нажмите x')]
                  ]
        window = sg.Window('Ошибка', layout)
        window.read()
    else:
        for i in list_emeil.keys():
            sg.one_line_progress_meter('Рассылка', progres_bar[progres_bar_chet + 1], len(list_emeil.keys()) * 2 + 3,'Формируем письмо')
            time.sleep(1)
            progres_bar_chet += 1
            main_send_email(list_emeil[i], subject,html_header_body(message_header, message_header_tabel, message_body_table, i, dolg[i]),sendler, password)
            sg.one_line_progress_meter('Рассылка', progres_bar[progres_bar_chet + 1], len(list_emeil.keys()) * 2 + 3,'Письмо отправлено')
            time.sleep(1)
            progres_bar_chet += 1





if __name__ == '__main__':
    print(main('timofei.kropachev@gmail.com','oqtspbctwzowkbfm','Main_new.xlsm')) #пароль используеться не от акаунта, а от Google преложения



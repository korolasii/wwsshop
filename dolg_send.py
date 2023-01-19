from send_emeil import main_send_email
from tabel_for_dolg import header_table_for_dolg
from jinja2 import Environment, FileSystemLoader
import PySimpleGUI as sg
import time



def html_header_body(message_header_tabel, buers, dolg):
    file_loader = FileSystemLoader('')
    env = Environment(loader=file_loader)

    tm = env.get_template('web/index_dolg.html')
    return tm.render(message=message_header_tabel,
                     buers=buers, dolg=dolg)




def dolg_send(sendler, password, way):
    dolg, list_emeil, progres_bar, message_for_header_table_dolg= header_table_for_dolg(way)
    progres_bar_chet = 2
    subject='Долги'
    for i in list_emeil.keys():
        sg.one_line_progress_meter('Рассылка', progres_bar[progres_bar_chet], len(list_emeil.keys()) * 2 + 2,'Формируем письмо')
        time.sleep(1)
        progres_bar_chet += 1
        main_send_email(list_emeil[i], subject,html_header_body(message_for_header_table_dolg, i, dolg[i]),sendler, password,)
        sg.one_line_progress_meter('Рассылка', progres_bar[progres_bar_chet], len(list_emeil.keys()) * 2 + 2,'Письмо отправлено')
        time.sleep(1)
        progres_bar_chet += 1

if __name__ == '__main__':
    print(dolg_send('timofei.kropachev@gmail.com','oqtspbctwzowkbfm','Main_new.xlsm'))
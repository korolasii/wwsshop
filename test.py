import eel
from main import main
import base_json_test
from dolg_send import dolg_send

eel.init('web')


@eel.expose
def recv_data(emeil, password):
    global info
    info = {
        'emeil': emeil,
        'password': password
    }
    return info

@eel.expose
def start_code(emeil, password):
    way='Main_new.xlsm'
    main(emeil, password, way)


@eel.expose
def reload():
    way = 'Main_new.xlsm'
    base_json_test.base_byer_reload(way)

@eel.expose
def dolg_distribution(emeil, password):
    way = 'Main_new.xlsm'
    dolg_send(emeil, password, way)



eel.start('index.html', size=(600, 600))

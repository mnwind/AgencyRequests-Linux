import PySimpleGUI as sg
from json import (load as jsonload, dump as jsondump)
from os import path



##################### Load/Save Settings File #####################
def load_settings(settings_file, default_settings):
    try:
        with open(settings_file, 'r') as f:
            settings = jsonload(f)
    except Exception as e:
        sg.popup_quick_message(f'exception {e}', 'Файл настроек не найден... будет создан со значениями по умолчанию', keep_on_top=True, background_color='red', text_color='white')
        settings = default_settings
        save_settings(settings_file, settings, None)
    return settings


def save_settings(settings_file, settings, values, SETTINGS_KEYS_TO_ELEMENT_KEYS):
    if values:      # if there are stuff specified by another window, fill in those values
        for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:  # update window with the values read from settings file
            try:
                settings[key] = values[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]]
            except Exception as e:
                print(f'Ошибка обновления настроек. Key = {key}')

    with open(settings_file, 'w') as f:
        jsondump(settings, f)

    sg.popup('Настройки сохранены')

#
def form():
    SETTINGS_FILE = path.join(path.dirname(__file__), r'settings_file.cfg')
    DEFAULT_SETTINGS = {'db_file': None , 'theme': sg.theme()}
    # "Map" from the settings dictionary keys to the window's element keys
    SETTINGS_KEYS_TO_ELEMENT_KEYS = {'db_file': '-DB FILE-' , 'theme': '-THEME-'}
    settings = load_settings(SETTINGS_FILE, DEFAULT_SETTINGS )
    c_theme = settings['theme']
    sg.theme(c_theme)

    def TextLabel(text): return sg.Text(text+':', justification='r', size=(20,1))

    frame_layout = [[sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Save_24x24.png'), key='-SAVE-', tooltip = 'Сохранить' ),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Log Out_24x24.png'), key='-EXIT-', tooltip = 'Выход' )]]

    layout = [  [sg.Text('Настройки')],
                [TextLabel('База данных'),sg.Input(key='-DB FILE-', default_text=settings['db_file']), sg.FileBrowse(target='-DB FILE-')],
                [TextLabel('Тема оформления'),sg.Combo(sg.theme_list(), size=(20, 20), key='-THEME-')],
                [sg.Frame('', frame_layout, element_justification = "center")]
            ]

    window = sg.Window('Настройки', layout, keep_on_top=False, finalize=True)



    while True:             # Event Loop

        for key in SETTINGS_KEYS_TO_ELEMENT_KEYS:   # update window with the values read from settings file
            try:
                window[SETTINGS_KEYS_TO_ELEMENT_KEYS[key]].update(value=settings[key])
            except Exception as e:
                print(f'Ошибка чтения настроек. Key = {key}')
        event, values = window.read()
        if event in (sg.WIN_CLOSED, '-EXIT-'):
            break
        if event == '-SAVE-':
            c_theme = values['-THEME-']
            sg.theme(c_theme)
            answ = sg.popup('Сохранить внесенные изменения ', custom_text=('Сохранить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Сохранить':
                save_settings(SETTINGS_FILE, settings, values, SETTINGS_KEYS_TO_ELEMENT_KEYS)

    window.close()

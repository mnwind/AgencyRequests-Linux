import PySimpleGUI as sg
import sqlite3 as sql
from os import path

def form (conn):
    agency_id = 1
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Agency_card WHERE id = "+ str(agency_id))
    results = cursor.fetchone()
    frame_layout = [[sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Save_24x24.png'), key='-SAVE-', tooltip = 'Сохранить' ),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Log Out_24x24.png'), key='-EXIT-', tooltip = 'Выход' )]]
    agencylayout = [
        [sg.T('Информация о предприятии', auto_size_text=True)],
        [sg.HorizontalSeparator()],
        [sg.T('Наименование', size=(15,1)), sg.In(results[1], size=(70,1))],
        [sg.T('Адрес', size=(15,1)), sg.In(results[2],size=(70,1))],
        [sg.T('ИНН', size=(15,1)), sg.In(results[3],size=(10,1)), sg.T('КПП', auto_size_text=True),
        sg.In(results[4],size=(9,1)), sg.T('ОГРН',auto_size_text=True), sg.In(results[5],size=(15,1)), sg.T('ОКВЭД',
        auto_size_text=True), sg.In(results[6],size=(5,1))],
        [sg.T('Телефон', size=(15,1)), sg.In(results[7],size=(12,1)), sg.T('E-mail', auto_size_text=True),
        sg.In(results[8],size=(15,1)), sg.T('WWW',auto_size_text=True), sg.In(results[9],size=(15,1))],
        [sg.T('Директор', size=(15,1)), sg.In(results[10],size=(25,1))],
        [sg.HorizontalSeparator()],
        [sg.T('Наименование банка', size=(15,1)), sg.In(results[11],size=(70,1))],
        [sg.T('Р/С', size=(15,1)), sg.In(results[12],size=(70,1))],
        [sg.T('К/С', size=(15,1)), sg.In(results[13],size=(70,1))],
        [sg.T('БИК', size=(15,1)), sg.In(results[14],size=(7,1))],
        [sg.HorizontalSeparator()],
        [sg.Frame('', frame_layout, element_justification = "center")]]

    agwnd = sg.Window('Информация о предприятии', agencylayout, no_titlebar=False)

    while True:
        event, values =agwnd.read()
        if event == '-EXIT-'  or event is None:
            break

        if event == '-SAVE-':
            answ = sg.popup('Сохранить внесенные изменения? ', custom_text=('Сохранить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Сохранить':
                column_values = list(values.values())
                upd_sql = "UPDATE Agency_card SET name = ?, adress = ?, inn = ?, kpp = ?, ogrn = ?, okved = ?, phone = ?, e_mail = ?, www = ?, boss = ?, bank_name  = ?, account = ?, cor_account = ?, bank_bik = ? WHERE id = " + str(agency_id) +";"
                cursor.execute(upd_sql, column_values)
                conn.commit()

    agwnd.close()

import PySimpleGUI as sg
import sqlite3 as sql
from os import path
import contformat

def updatewnd(townd, results1):     # Обновление правой колонки макета

    townd['-FNAME-'].update(results1[1])
    townd['-SNAME-'].update(results1[2])
    townd['-ADR-'].update(results1[3])
    townd['-INN-'].update(results1[4])
    townd['-KPP-'].update(results1[5])
    townd['-TEL-'].update(results1[6])
    townd['-EMAIL-'].update(results1[7])
    townd['-NREE-'].update(results1[8])
    townd['-SITE-'].update(results1[9])
    townd['-STRA-'].update(results1[10])
    townd['-ADST-'].update(results1[11])
    townd['-STTEL-'].update(results1[12])
    townd['-STDO-'].update(results1[13])
    townd['-DATEB-'].update(results1[14])
    townd['-DATEE-'].update(results1[15])

def updatelst(conn, townd):       # Обновление списка туроператоров
    cursor = conn.cursor()
    cursor.execute("SELECT name_short_to FROM Cat_tourop ORDER BY name_short_to ASC")
    results = cursor.fetchall()
    if results == []:
        sg.popup('Список туроператоров пуст. Будет вставлена пустая запись')
        ins_sql = "INSERT INTO Cat_tourop (name_short_to) VALUES (' ');"
        cursor.execute(ins_sql)
        conn.commit()
        cursor.execute("SELECT name_short_to FROM Cat_tourop ORDER BY name_short_to ASC")
        results = cursor.fetchall()
    townd['-LIST-'].update(results)
    return results[0]

def form (conn):
#   работа с таблицей БД по туроператорам
#   возвращает ID оператора или 0 если выход без выбора
#
#   получение списка туроператоров
    cursor = conn.cursor()
    cursor.execute("SELECT name_short_to FROM Cat_tourop ORDER BY name_short_to ASC")
    results = cursor.fetchall()
    if results == []:
        sg.popup('Список туроператоров пуст. Будет вставлена пустая запись')
        ins_sql = "INSERT INTO Cat_tourop (name_short_to) VALUES (' ');"
        cursor.execute(ins_sql)
        conn.commit()
        cursor.execute("SELECT name_short_to FROM Cat_tourop ORDER BY name_short_to ASC")
        results = cursor.fetchall()
#   получение информации по первому туроператору
    s_name = results[0]
    cursor.execute("SELECT * FROM Cat_tourop WHERE name_short_to=?",s_name)
    results1 = cursor.fetchone()
#   макет окна
    frame_layout = [[sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Save_24x24.png'), key='-SAVE-', tooltip = 'Сохранить' ),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Log Out_24x24.png'), key='-EXIT-', tooltip = 'Выход' )]]
    frame1_layout = [[ sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Add_24x24.png'), key='-ATO-', tooltip = 'Новый туроператор'),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Delete_24x24.png'), key='-DTO-', tooltip = 'Удалить'),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Check_24x24.png'), key='-CTO-', tooltip = 'Выбор туроператора для договора')]]
    column_to = [[sg.Text('Информация о туроператоре')],
    [sg.T('Полное наименование', size=(20,1)), sg.In(size=(70,1), key='-FNAME-', default_text=results1[1])],
    [sg.T('Адрес', size=(20,1)), sg.In(size=(70,1), key='-ADR-', default_text=results1[3])],
    [sg.T('Сокращенное наименование', size=(20,1)), sg.In(size=(30,1), key='-SNAME-', default_text=results1[2])],
    [sg.T('ИНН', size=(20,1)), sg.In(size=(10,1), key='-INN-', default_text=results1[4]), sg.T('КПП', auto_size_text=True),
    sg.In(size=(9,1), key='-KPP-', default_text=results1[5]), sg.T('Номер в реестре', auto_size_text=True), sg.In(size=(10,1), key='-NREE-',default_text=results1[8])],
    [sg.T('Телефон', size=(20,1)), sg.In(size=(12,1), key='-TEL-', default_text=results1[6]),
    sg.T('E-mail', auto_size_text=True), sg.In(size=(20,1), key='-EMAIL-', default_text=results1[7]),
    sg.T('Сайт',auto_size_text=True), sg.In(size=(20,1), key='-SITE-', default_text=results1[9])],
    [sg.HorizontalSeparator()],
    [sg.T('Финансовое обеспечение', size=(40,1))],
    [sg.T('Наименование страховщика', size=(20,1)), sg.In(size=(70,1), key='-STRA-', default_text=results1[10])],
    [sg.T('Адрес', size=(20,1)), sg.In(size=(70,1), key='-ADST-', default_text=results1[11])],
    [sg.T('Наименование договора', size=(20,1)), sg.In(size=(70,1), key='-STDO-', default_text=results1[13])],
    [sg.T('Период действия, с:', size=(20,1)), sg.In(size=(10,1), key='-DATEB-', default_text=results1[14]),
    sg.T('по:', auto_size_text=True), sg.In(size=(10,1), key='-DATEE-', default_text=results1[15]),
    sg.T('Телефон', auto_size_text=True), sg.In(size=(12,1), key='-STTEL-', default_text=results1[12])],
    [sg.HorizontalSeparator()],
    [sg.Frame('', frame_layout, element_justification = "center")]]

    column_to_list = [[sg.Text('Список туроператоров')],
    [sg.Listbox(values=results, size=(30, 15), key='-LIST-', default_values=results[0], enable_events=True, auto_size_text=True, pad=(5, 5), select_mode=sg.LISTBOX_SELECT_MODE_SINGLE)],
    [sg.HorizontalSeparator()],
    [sg.Frame('', frame1_layout, element_justification = "center")]]

    tolayout = [[ sg.Column(column_to_list), sg.Column(column_to)]]

    townd = sg.Window('Список туроператоров', tolayout, no_titlebar=False)

    while True:
        event, values =townd.read()
        if event == "-LIST-":   #перемещение по списку
            p1 = values['-LIST-']
            s_name = p1[0]
            cursor.execute("SELECT * FROM Cat_tourop WHERE name_short_to=?",s_name)
            results1 = cursor.fetchone()
            updatewnd(townd, results1)

        if event == '-EXIT-'  or event is None:
            id_oper = 0
            break

        if event == '-ATO-':
            answ = sg.popup_get_text('Введите сокращенное имя нового оператора')
            if answ != None:
#               вставка записи с введенным именем
                ins_sql = "INSERT INTO Cat_tourop (name_short_to) VALUES ('" + answ + "');"
                cursor.execute(ins_sql)
                conn.commit()
                s_name = updatelst(conn, townd)
#               получение данных по умолчанию
                sel_sql = "SELECT * FROM Cat_tourop WHERE name_short_to='"+ answ + "';"
                cursor.execute(sel_sql)
                results1 = cursor.fetchone()
                updatewnd(townd, results1)

        if event == '-DTO-':
            answ = sg.popup('Удалить данные оператора ' + results1[2] + '. В том числе из существующих заявок?', custom_text=('Удалить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Удалить':
#               замена на 0 id туроператора в заявках
                upd_sql = "UPDATE list_request SET id_to = 0 WHERE id_to = " + str(results1[0])
                cursor.execute(upd_sql)
                conn.commit()
#               удаление записи по туроператору
                del_sql = "DELETE FROM Cat_tourop WHERE name_short_to = '" + results1[2] + "';"
                cursor.execute(del_sql)
                conn.commit()
#               обновление списка
                s_name = updatelst(conn, townd)
#               получение информации по первому в списке и обновление правой части
                sel_sql = "SELECT * FROM Cat_tourop WHERE name_short_to='"+ s_name[0] + "';"
                cursor.execute(sel_sql)
                results1 = cursor.fetchone()
                updatewnd(townd, results1)
        if event == '-CTO-':
            id_oper = results1[0]
            break

        if event == '-SAVE-':
            answ = sg.popup('Сохранить внесенные изменения ' + results1[2], custom_text=('Сохранить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Сохранить':
                if contformat.contdatef(values['-DATEB-']) and contformat.contdatef(values['-DATEE-']):
                    upd_sql = "UPDATE Cat_tourop SET name_full_to = '" + str(values['-FNAME-']) + "', name_short_to = '" + str(values['-SNAME-']) + "', adress_to = '" + str(values['-ADR-']) + "', inn_to = '" + str(values['-INN-']) + "', kpp_to = '" + str(values['-KPP-']) + "', tel_to = '" + str(values['-TEL-']) + "', email_to = '" + str(values['-EMAIL-']) + "', num_fedr_to = '" + str(values['-NREE-']) + "', site = '" + str(values['-SITE-']) + "', name_strah = '" + str(values['-STRA-']) + "', adress_strah = '" + str(values['-ADST-']) + "', tel_strah = '" + str(values['-STTEL-']) + "', text_strah = '" + str(values['-STDO-']) + "', date_beg_strah = '" + str(values['-DATEB-']) + "', date_end_strah = '" + str(values['-DATEE-']) + "' WHERE id_to = '" + str(results1[0]) + "';"
                    cursor.execute(upd_sql)
                    conn.commit()
                    s_name = updatelst(conn, townd)
                else:
                    sg.popup('Ошибка в формате даты (правильный формат "дд.мм.гггг")')

    townd.close()

    return id_oper

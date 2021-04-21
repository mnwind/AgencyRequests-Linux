import PySimpleGUI as sg
import sqlite3 as sql
import contformat
from os import path

def updatewnd(cuwnd, results1):     # Обновление правой колонки макета
    cuwnd['-FIO-'].update(results1[1])
    cuwnd['-ADR-'].update(results1[2])
    cuwnd['-LPASS-'].update(results1[3])
    cuwnd['-DLPASS-'].update(results1[4])
    cuwnd['-WLPASS-'].update(results1[5])
    cuwnd['-FPASS-'].update(results1[6])
    cuwnd['-DIFP-'].update(results1[7])
    cuwnd['-DEFP-'].update(results1[8])
    cuwnd['-WFPASS-'].update(results1[9])
    cuwnd['-FNCUST-'].update(results1[10])
    cuwnd['-FLCUST-'].update(results1[11])
    cuwnd['-CTEL-'].update(results1[12])
    cuwnd['-CMAIL-'].update(results1[13])
    cuwnd['-DATER-'].update(results1[14])

def updatelst(conn, cuwnd):       # Обновление списка клиентов
    cursor = conn.cursor()
    cursor.execute("SELECT fio FROM Cat_cust ORDER BY fio ASC")
    results = cursor.fetchall()
    if results == []:
        sg.popup('Список клиентов пуст. Будет вставлена пустая запись')
        ins_sql = "INSERT INTO Cat_cust (fio) VALUES (' ');"
        cursor.execute(ins_sql)
        conn.commit()
        cursor.execute("SELECT fio FROM Cat_cust ORDER BY fio ASC")
        results = cursor.fetchall()
    cuwnd['-LIST-'].update(results)
    return results[0]

def form (conn):
#   работа с таблицей БД по клиентам
#   возвращает ID клиента или 0 если выход без выбора
#
# Получение списка клиентов
    cursor = conn.cursor()
    cursor.execute("SELECT fio FROM Cat_cust ORDER BY fio ASC")
    results = cursor.fetchall()
    if results == []:
        sg.popup('Список клиентов пуст. Будет вставлена пустая запись')
        ins_sql = "INSERT INTO Cat_cust (fio) VALUES (' ');"
        cursor.execute(ins_sql)
        conn.commit()
# Получение данных по первому клиенту
    s_name = results[0]
    cursor.execute("SELECT * FROM Cat_cust WHERE fio=?",s_name)
    results1 = cursor.fetchone()
# Макет окна
    frame_layout = [[sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Save_24x24.png'), key='-SAVE-', tooltip = 'Сохранить' ),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Log Out_24x24.png'), key='-EXIT-', tooltip = 'Выход' )]]
    frame1_layout = [[ sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Add_24x24.png'), key='-ACU-', tooltip = 'Новый турист'),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Delete_24x24.png'), key='-DCU-', tooltip = 'Удалить'),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Check_24x24.png'), key='-CCU-', tooltip = 'Выбор туриста для договора')]]
    column_to = [[sg.Text('Информация о заказчике')],
    [sg.T('ФИО', size=(15,1)), sg.In(size=(70,1), key='-FIO-', default_text=results1[1])],
    [sg.T('Адрес', size=(15,1)), sg.In(size=(70,1), key='-ADR-', default_text=results1[2])],
    [sg.T('Дата рождения', size=(15,1)), sg.In(size=(10,1), key='-DATER-', default_text=results1[14]),
    sg.T('Телефон', auto_size_text=True), sg.In(size=(14,1), key='-CTEL-', default_text=results1[12]),
    sg.T('E-mail', auto_size_text=True), sg.In(size=(20,1), key='-CMAIL-', default_text=results1[13])],
    [sg.HorizontalSeparator()],
    [sg.T('Данные паспорта РФ', size=(40,1))],
    [sg.T('Серия Номер', size=(15,1)), sg.In(size=(30,1), key='-LPASS-', default_text=results1[3]),
    sg.T('Дата выдачи', auto_size_text=True), sg.In(size=(10,1), key='-DLPASS-', default_text=results1[4])],
    [sg.T('Кем выдан', size=(15,1)), sg.In(size=(70,1), key='-WLPASS-', default_text=results1[5])],
    [sg.HorizontalSeparator()],
    [sg.T('Данные заграничного паспорта', size=(40,1))],
    [sg.T('Серия Номер', size=(15,1)), sg.In(size=(10,1), key='-FPASS-', default_text=results1[6]),
    sg.T('Действителен c', auto_size_text=True), sg.In(size=(10,1), key='-DIFP-', default_text=results1[7]),
    sg.T('по',auto_size_text=True), sg.In(size=(10,1), key='-DEFP-', default_text=results1[8])],
    [sg.T('Подразделение', size=(15,1)), sg.In(size=(20,1), key='-WFPASS-', default_text=results1[9])],
    [sg.T('Имя (LAT)', size=(15,1)), sg.In(size=(15,1), key='-FNCUST-', default_text=results1[10]),
    sg.T('Фамилия (LAT)', auto_size_text=True), sg.In(size=(35,1), key='-FLCUST-', default_text=results1[11])],
    [sg.HorizontalSeparator()],
    [sg.Frame('', frame_layout, element_justification = "center")]]

    column_to_list = [[sg.T('Список Заказчиков')],
    [sg.T('Фильтр', size=(8,1)), sg.In(size=(20, 1), enable_events=True, key='-FILTR-')],
    [sg.Listbox(values=results, size=(30, 15), key='-LIST-', enable_events=True, auto_size_text=True, pad=(5, 5), select_mode=sg.LISTBOX_SELECT_MODE_SINGLE)],
    [sg.HorizontalSeparator()],
    [sg.Frame('', frame1_layout, element_justification = "center")]]

    tolayout = [[ sg.Column(column_to_list), sg.Column(column_to)]]

    cuwnd = sg.Window('Список заказчиков', tolayout, margins = (15, 15), no_titlebar = False, border_depth = 5)

    while True:     # Обработка событий
        event, values =cuwnd.read()
        if event == "-LIST-":       # Перемещение по списку
            p1 = values['-LIST-']
            s_name = p1[0]
            cursor.execute("SELECT * FROM Cat_cust WHERE fio=?",s_name)
            results1 = cursor.fetchone()
            updatewnd(cuwnd, results1)

        if event == '-EXIT-'  or event is None:
            id_cust = 0
            break

        if event == '-ACU-':
            answ = sg.popup_get_text('Введите ФИО нового заказчика')
            if answ != None:
                ins_sql = "INSERT INTO Cat_cust (fio) VALUES ('" + answ + "');" #Вставка записи с введенным именем
                cursor.execute(ins_sql)
                conn.commit()
                s_name = updatelst(conn, cuwnd)   # Обновление списка
                sel_sql = "SELECT * FROM Cat_cust WHERE fio='"+ answ + "';" # Получение данных по умолчанию
                cursor.execute(sel_sql)
                results1 = cursor.fetchone()
                updatewnd(cuwnd, results1)  # Обновление правой колонки

        if event == '-DCU-':
            answ = sg.popup('Удалить данные по заказчику ' + results1[1] + '. В том числе из существующих заявок?', custom_text=('Удалить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Удалить':
#               замена на 0 id заказчика в списке заявок
                upd_sql = "UPDATE list_request SET id_cust = 0 WHERE id_cust = " + str(results1[0])
                cursor.execute(upd_sql)
                conn.commit()
#               удаление записей по id клиента из списка туристов по заявке
                del_sql = "DELETE FROM req_tourist WHERE id_cust = " + str(results1[0])
                cursor.execute(del_sql)
                conn.commit()
#               Удаление данных по клиенту
                del_sql = "DELETE FROM Cat_cust WHERE fio = '" + results1[1] + "';"
                cursor.execute(del_sql)
                conn.commit()
                s_name = updatelst(conn, cuwnd)   # Получение списка клиентов
                sel_sql = "SELECT * FROM Cat_cust WHERE fio='"+ s_name[0] + "';"
                cursor.execute(sel_sql)
                results1 = cursor.fetchone()
                updatewnd(cuwnd, results1)  # Обновление данных по первому в списке

        if event == '-CCU-':
            id_cust = results1[0]
            break

        if event == '-SAVE-':
            answ = sg.popup('Сохранить внесенные изменения ' + results1[1], custom_text=('Сохранить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Сохранить':
                if contformat.contdatef(values['-DATER-']) and contformat.contdatef(values['-DLPASS-']) and contformat.contdatef(values['-DIFP-']) and contformat.contdatef(values['-DEFP-']):
                    upd_sql = "UPDATE Cat_cust SET fio = '" + str(values['-FIO-']) + "', cust_adress = '" + str(values['-ADR-']) + "', num_local_pass = '" + str(values['-LPASS-']) + "', date_local_pass = '" + str(values['-DLPASS-']) + "', who_local_pass = '" + str(values['-WLPASS-']) + "', num_for_pass = '" + str(values['-FPASS-']) + "', date_iss_for_pass = '" + str(values['-DIFP-']) + "', date_end_for_pass = '" + str(values['-DEFP-']) + "', who_for_pass = '" + str(values['-WFPASS-']) + "',first_name = '" + str(values['-FNCUST-']) + "', last_name = '" + str(values['-FLCUST-']) + "', cust_tel = '" + str(values['-CTEL-']) + "', cust_email = '" + str(values['-CMAIL-']) + "', date_r = '" + str(values['-DATER-']) + "' WHERE id_cust = '" + str(results1[0]) + "';"
                    cursor.execute(upd_sql) # Запись в БД
                    conn.commit()
                    s_name = updatelst(conn, cuwnd)   #Обновление списка (возможно изменение имени)
                else:
                    sg.popup('Ошибка в формате даты (правильный формат "дд.мм.гггг")')

        if values['-FILTR-'] != '':                         # фильтр не пуст
            search = values['-FILTR-']
            cursor.execute("SELECT fio FROM Cat_cust ORDER BY fio ASC") # Получение последнего списка
            results = cursor.fetchall()
            new_values = [x for x in results if search in x[0]]  # Фильтрация
            cuwnd['-LIST-'].update(new_values)      # Новый список
        else:
            s_name = updatelst(conn, cuwnd)       # Восстановление списка
    cuwnd.close()

    return id_cust

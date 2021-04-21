import PySimpleGUI as sg
import sqlite3 as sql
import custform # подумать как избежать двух импортов
import touropform
from os import path
import subprocess
from datetime import date
from docxtpl import DocxTemplate
import contformat

def dogform(cursor, results, cust_id, id_oper, req_id):
# формирование договора из шаблона  
    doc_filename=path.join('tpl', 'contracttpl.docx')
    doc = DocxTemplate(doc_filename)
#   формирование строки передаваемой в редактор   
#   блок полей заказчика    
    if cust_id == 0:
        results1 = ['', '', '', '', '', '', '']
    else:
        cursor.execute("SELECT fio, num_local_pass, date_local_pass, who_local_pass, cust_adress, cust_tel, cust_email \
        FROM Cat_cust WHERE id_cust = "+ str(cust_id))
        results1 = cursor.fetchone()
#   блок полей оператора
    if id_oper == 0:
        results2 = ['', '', '', '', '', '', '', '', '', '', '', '', '']
    else:
        cursor.execute("SELECT name_full_to, name_short_to, adress_to, tel_to, site, email_to, num_fedr_to, text_strah,\
        date_beg_strah, date_end_strah, name_strah, adress_strah, tel_strah FROM cat_tourop WHERE id_to = "+ str(id_oper))
        results2 = cursor.fetchone()
#   блок полей отелей
    hotel_contents = []
    cursor.execute("SELECT hotel, hotel_addr, quant_room, type_room, accom, date_begin, date_end, meal FROM req_accom WHERE id_req = "+ str(req_id))
    while True:
        results3 = cursor.fetchone()
        if results3 == None:
            break
        hotel_contents.append({'name_hotel' : results3[0], 'adr_hotel' : results3[1], 'q_room' : results3[2], \
        't_room' : results3[3], 'accom' : results3[4], 'dateb' : results3[5], 'datee' : results3[6], 'meal' : results3[7]},)
#   блок полей туристов
    tourist_contents = []
    cursor.execute("SELECT id_cust FROM req_tourist WHERE id_req = '" +str(req_id) +"'")
    rows = cursor.fetchall()
    for row in rows:
        cursor.execute("SELECT fio, date_r, num_for_pass, date_iss_for_pass, date_end_for_pass FROM cat_cust WHERE id_cust = "+ str(row[0]))
        results3 = cursor.fetchone()
        tourist_contents.append({'fio' : results3[0], 'dater' : results3[1], 'numpass' : results3[2], 'dateiss' : results3[3], \
        'dateendpass' : results3[4]},)
#   блок полей перевозки
    trans_contents = []
    cursor.execute("SELECT type_trans, route, date_there, date_back FROM req_trans WHERE id_req = "+ str(req_id))
    while True:
        results5 = cursor.fetchone()
        if results5 == None:
            break
        trans_contents.append({'typetrans' : results5[0], 'routetrans' : results5[1], 'dthere' : results5[2], \
        'dback' : results5[3],},)
#   данные агентства
    cursor.execute("SELECT * FROM Agency_card")
    results6 = cursor.fetchone()

    context =   {'trans_contents' : trans_contents, 'trst_contents' : tourist_contents, 'htl_contents' : hotel_contents, 'customer' : results1[0], 'custpass' : results1[1], 'custpassdate' : results1[2],\
    'custpasswho' : results1[3], 'custadr' : results1[4], 'custtel' : results1[5], 'custemail' : results1[6], \
    'namefullto' : results2[0], 'nameshortto' : results2[1], 'adressto' : results2[2], 'tellto' : results2[3], 'siteto' : results2[4], \
    'emailto' : results2[5], 'numreestr' : results2[6], 'textstrah' : results2[7], 'datebegstrah' : results2[8], \
    'dateendstrah' : results2[9], 'namestrah' : results2[10], 'adressstrah' : results2[11], 'telstrah' : results2[12], \
    'datereq' : results[1], 'numbcontr' : results[2], 'country' : results[3], 'region' : results[4], 'datetour' : results[6], \
    'dateendtour' : results[7], 'quantn' : results[8], 'ticket' : results[9], 'transfer' : results[10], 'excurprog' : results[11], \
    'otherserv' : results[12], 'tourgid' : results[13], 'transl' : results[14], 'lead' : results[15], 'viza' : results[16], \
    'med' : results[17], 'acc' : results[18], 'fail' : results[19], 'datefullpay' : results[24], 'costrub' : results[29], \
    'costcur' : results[28], 'curtour' : results[21], 'prepayrub' : results[30],'datedoc' : results[26], 'rateto' : results[31], \
    'nameag' : results6[1], 'adressag' : results6[2], 'innag' : results6[3], 'kppag' : results6[4], 'ogrnag' : results6[5], \
    'okvedag' : results6[6], 'phoneag' : results6[7], 'emailag' : results6[8], 'siteag' : results6[9], 'bossag' : results6[10], \
    'bankag' : results6[11], 'accountag' : results6[12], 'coraccount' : results6[13], 'bikag' : results6[14]}
    doc.render(context)
    gen_doc_filename = path.join('tpl', 'contractgen.docx')
    try:
        doc.save(gen_doc_filename)
        subprocess.call(["xdg-open", gen_doc_filename])
    except:
        sg.popup('Закройте временный файл')

def listcust(cursor,req_id):
# получение информации из списка туристов по заявке
    results4 = []
    listidtour_sql = "SELECT id_cust FROM req_tourist WHERE id_req = '" +str(req_id) +"'"
    cursor.execute(listidtour_sql)
    rows = cursor.fetchall()
    for row in rows:
        listtourist_sql = "SELECT id_cust, fio, date_r, num_for_pass, date_iss_for_pass, date_end_for_pass, who_for_pass FROM Cat_cust WHERE id_cust = '" + str(row[0]) +"'"
        cursor.execute(listtourist_sql)
        r4 = cursor.fetchone()
        results4.append(r4)
    if results4==[]:
        results4 = [('','', '','', '', '', '')]
    return results4

def listhotel(cursor, req_id):
#   получение информации из таблицы размещений по ИД заявки
    cursor.execute("SELECT no_in_table, date_begin, date_end, hotel, type_room, quant_room, accom, meal, hotel_addr FROM req_accom WHERE id_req = "+ str(req_id))
    results3 = cursor.fetchall()
#   если гостиницы нет то создать пустой список
    if results3==[]:
        results3 = [('', '', '', '', '', '', '', '','')]
    return results3

def reqhotelform(conn,results3,req_id):
#   форма редактирования отеля
    reqhotellayout = [
    [sg.T('Проживание с:',size=(16,1)), sg.In(results3[1], size=(10,1), key='-DATEHB-'), sg.T('по:', auto_size_text=True),
    sg.In(results3[2], size=(10,1), key='-DATEHE-'), sg.T('Отель ', auto_size_text=True), sg.In(results3[3], size=(30,1), key='-REQHOTEL-')],
    [sg.T('Номер ', auto_size_text=True), sg.In(results3[4], size=(20,1), key='-NROOMHOTEL-'),sg.T('Количество номеров ', auto_size_text=True), 
    sg.In(results3[5], size=(2,1), key='-QROOMHOTEL-'), sg.T('Тип размещения ', auto_size_text=True),
    sg.In(results3[6], size=(20,1), key='-TROOMHOTEL-'), sg.T('Питание ', auto_size_text=True), sg.In(results3[7], 
    size=(15,1), key='-MEALHOTEL-')],
    [sg.T('Адрес отеля ', auto_size_text=True), sg.In(results3[8], size=(50,1), key='-ADRHOTEL-')],
    [sg.Button('Сохранить'), sg.Button('Выход')]
    ]
    rhwnd = sg.Window('Отели по туру', reqhotellayout, no_titlebar=False)
    while True:
        event, values =rhwnd.read()
        if event == 'Выход'  or event is None:
            break
        if event in ('Сохранить'):
            if contformat.contdatef(values['-DATEHB-']) and contformat.contdatef(values['-DATEHE-']):
                column_values = (values['-DATEHB-'], values['-DATEHE-'], values['-REQHOTEL-'], values['-ADRHOTEL-'], values['-NROOMHOTEL-'], values['-QROOMHOTEL-'], values['-TROOMHOTEL-'], values['-MEALHOTEL-'], req_id, results3[0])
                upd_sql = "UPDATE req_accom SET date_begin = ?, date_end = ?, hotel = ?, hotel_addr = ?, type_room = ?, quant_room = ?, accom = ?, meal = ? WHERE id_req = ? AND no_in_table= ?;"
                cursor = conn.cursor()
                cursor.execute(upd_sql,column_values)
                conn.commit()
                break
            else:
                sg.popup('Ошибка в формате даты (правильный формат "дд.мм.гггг")')

    rhwnd.close()

def listtrans(cursor,req_id):
#   список транспорта по заявке
#   получение информации из таблицы транспорта по ИД заявки
    cursor.execute("SELECT no_in_trans, type_trans, route, date_there, date_back FROM req_trans WHERE id_req = "+ str(req_id) + " ORDER BY no_in_trans ASC")
    results5 = cursor.fetchall()
#   если транспорта нет то создать пустой список
    if results5==[]:
        results5 = [('', '', '', '', '')]
    return results5

def reqtransform(conn, results5, req_id):
#   форма редактирования отеля
    reqtranslayout = [
    [sg.T('Тип перевозки', auto_size_text=True), sg.Combo(('Авиабилет','Ж.д.билет','Трансфер'), default_value=results5[1], size=(10,1), key='-TYPETRANS-'), 
    sg.T('Маршрут', auto_size_text=True), sg.In(results5[2], size=(50,1), key='-ROUTE-'), 
    sg.T('Дата туда ', auto_size_text=True), sg.In(results5[3], size=(10,1), key='-DATET-'),
    sg.T('Дата обратно ', auto_size_text=True), sg.In(results5[4], size=(10,1), key='-DATEBACK-')],
    [sg.Button('Сохранить'), sg.Button('Выход')]
    ]
    rtwnd = sg.Window('Перевозки по туру', reqtranslayout, no_titlebar=False)
    while True:
        event, values =rtwnd.read()
        if event == 'Выход'  or event is None:
            break
        if event in ('Сохранить'):
            if contformat.contdatef(values['-DATET-']) and contformat.contdatef(values['-DATEBACK-']):
                column_values = (values['-TYPETRANS-'], values['-ROUTE-'], values['-DATET-'], values['-DATEBACK-'], req_id, results5[0])
                upd_sql = "UPDATE req_trans SET type_trans = ?, route = ?, date_there = ?, date_back = ? WHERE id_req = ? AND no_in_trans= ?;"
                cursor = conn.cursor()
                cursor.execute(upd_sql,column_values)
                conn.commit()
                break
            else:
                sg.popup('Ошибка в формате даты (правильный формат "дд.мм.гггг")')
    rtwnd.close()


def form (conn, req_id):
    cursor = conn.cursor()
    if req_id == 0: #если заявка новая создается запись, получается id
        cursor.execute("INSERT INTO list_request (date_req) VALUES ('" + str(date.today().strftime("%d.%m.%Y")) + "')")
        conn.commit()
        cursor.execute("SELECT id_req FROM list_request WHERE rowid=last_insert_rowid()")
        r = cursor.fetchone()
        req_id = r[0]
 
#   получение информации по заявке
    cursor.execute("SELECT * FROM list_request WHERE id_req = "+ str(req_id))
    results = cursor.fetchone()

#   получение информации по ИД клиента
    if results[5] == 0:
        results1 = ['']
        cust_id = 0
    else:
        cursor.execute("SELECT fio FROM Cat_cust WHERE id_cust = "+ str(results[5]))
        results1 = cursor.fetchone()
        cust_id = results[5]

#   получение информации по ИД туроператора
    if results[20] == 0:
        results2 = ['']
        id_oper = 0
    else:
        cursor.execute("SELECT name_short_to FROM Cat_tourop WHERE id_to = "+ str(results[20]))
        results2 = cursor.fetchone()
        id_oper = results[20]

#   получение информации из таблицы размещений по ИД заявки
    results3 = listhotel(cursor, req_id)
#   Заголовок для таблицы гостиниц в форме
    header_list_hotel = ['№ ', 'Заезд', 'Выезд', 'Отель', 'Номер', 'К-во', 'Размещение', 'Питание', 'Адрес отеля' ]
#   получение информации из таблицы транспорта по ИД заявки
    results5 = listtrans(cursor, req_id)
    header_list_trans = ['№', 'Тип', 'Маршрут (вид, класс)', 'Дата туда', 'Дата обратно']
#   получение информации из таблицы туристов по ИД заявки
    results4 = listcust(cursor,req_id)
    header_list_turists = ['ID','Фамилия Имя Отчество', 'Дата рождения', 'Номер З.П.', 'Дата выдачи', 'Действует по', 'Подр.' ]
#Макет окна заявки
    frame_layout = [[sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Text Document_24x24.png'), key='-DOGV-', tooltip = 'Договор' ),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Save_24x24.png'), key='-SAVE-', tooltip = 'Сохранить' ),
    sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Log Out_24x24.png'), key='-EXIT-', tooltip = 'Выход' )]]
    agencylayout = [
        [sg.T('Заявка №', auto_size_text=True), sg.T(text = str(results[0]), size=(4,1)), sg.T('от', auto_size_text=True),
        sg.In(results[1], size=(10,1), key='-DATR-'), sg.CalendarButton(button_text='', image_filename=path.join('ico', 'Calendar_24x24.png'), target='-DATR-', format='%d.%m.%Y'),
        sg.T('к договору №', auto_size_text=True),  sg.In(results[2], size=(9,1), key='-NCONTR-'),
        sg.T('Заказчик', auto_size_text=True), sg.T(text = str(results1[0]), size=(30,1), relief=sg.RELIEF_SUNKEN, key='-FIO-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'User_24x24.png'), key='-CUST-', border_width=0)],
        [sg.T('Страна', size=(6,1)), sg.In(results[3], size=(20,1)), sg.T('Регион', auto_size_text=True), sg.In(results[4], size=(20,1)),
        sg.T('с', auto_size_text=True), sg.In(results[6], size=(10,1), key='-DATEB-'),
#        sg.CalendarButton(button_text='', image_filename=path.join('ico', 'Calendar_24x24.png'), target='-DATEB-', format='%d.%m.%Y'),
        sg.T('по', auto_size_text=True), sg.In(results[7], size=(10,1), key='-DATEE-'),
#        sg.CalendarButton(button_text='', image_filename=path.join('ico', 'Calendar_24x24.png'), target='-DATEE-', format='%d.%m.%Y'),
        sg.T('ночей', auto_size_text=True), sg.In(results[8], size=(2,1), key='-NNIGHT-')],
        [sg.T('Билет', size=(6,1), visible=False), sg.In(results[9],visible=False, size=(50,1)), sg.T('Трансфер', visible=False, auto_size_text=True), sg.In(results[10], visible=False, size=(30,1))],
        [sg.HorizontalSeparator()],
        [sg.T('Транспорт:' , size=(9,1)),
        sg.Table( values=results5 , headings=header_list_trans, num_rows=2, key='-LTRANS-', enable_events=True, pad=(1, 1), auto_size_columns=False, col_widths=[2, 10, 48, 10, 10], select_mode=sg.TABLE_SELECT_MODE_BROWSE)],
        [sg.T('', size=(9,1)), sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Add_24x24.png'), key='-ATRANS-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Properties_24x24.png'), key='-MTRANS-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Delete_24x24.png'), key='-DTRANS-')],
        [sg.HorizontalSeparator()],
        [sg.T('Отели:' , size=(6,1)),
        sg.Table( values=results3 , headings=header_list_hotel, num_rows=1, key='-LHOTEL-', enable_events=True, pad=(1, 1), select_mode=sg.TABLE_SELECT_MODE_BROWSE, auto_size_columns=False, col_widths=[2, 10, 10, 20,8, 3, 10, 10, 10])],
        [sg.T('', size=(6,1)), sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Add_24x24.png'), key='-AHOTEL-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Properties_24x24.png'), key='-MHOTEL-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Delete_24x24.png'), key='-DHOTEL-')],
        [sg.HorizontalSeparator()],
        [sg.T('Экскурсионная программа', auto_size_text=True), sg.In(results[11],size=(30,1), key='-EPROG-'),
        sg.T('Прочие услуги', auto_size_text=True), sg.In(results[12],size=(40,1), key='-PRUS-')],
        [sg.T('Гид', size=(20,1)), sg.Combo(('Да','Нет'), default_value=results[13], size=(3,1), key='-GID-'),
        sg.T('Экскурсовод', auto_size_text=True), sg.Combo(('Да','Нет'), default_value=results[14], size=(3,1), key='-EXCUR-'),
        sg.T('Руководитель группы', auto_size_text=True), sg.Combo(('Да','Нет'), default_value=results[15], size=(3,1), key='-LEAD-'),
        sg.T('Виза', auto_size_text=True), sg.Combo(('Да','Нет'), default_value=results[16], size=(3,1), key='-VISA-')],
        [sg.HorizontalSeparator()],
        [sg.T('Страхование: медицинское', auto_size_text=True), sg.Combo(('Да','Нет'), default_value=results[17], size=(3,1), key='-MEDS-'),
        sg.T('от несчасного случая', auto_size_text=True), sg.Combo(('Да','Нет'), default_value=results[18], size=(3,1), key='-NSS-'),
        sg.T('от невыезда', auto_size_text=True), sg.Combo(('Да','Нет'), default_value=results[19], size=(3,1), key='-NEVS-'),
        sg.T('примечание', auto_size_text=True), sg.In(results[32], size=(20,1), key='-PRIM-')],
        [sg.HorizontalSeparator()],
        [sg.T('Туристы:' , size=(6,1)),
        sg.Table( values=results4 , headings=header_list_turists, num_rows=4, key='-LTURISTS-', enable_events=True, pad=(5, 5), select_mode=sg.TABLE_SELECT_MODE_BROWSE, auto_size_columns=False, col_widths=[3, 30, 10, 10, 10, 10, 10])],
        [sg.T('', size=(6,1)), sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Add_24x24.png'), key='-ATURIST-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Delete_24x24.png'), key='-DTURIST-')],
        [sg.HorizontalSeparator()],
        [sg.T('Оператор', auto_size_text=True), sg.T(text = str(results2[0]), relief=sg.RELIEF_SUNKEN, size=(30,1), key = '-TO-'),
        sg.Button('', auto_size_button=True, image_filename=path.join('ico', 'Globe_24x24.png'), key='-TOUROP-', border_width=0),
        sg.T(('Номер заявки у туроператора'), auto_size_text=True), sg.In(results[33], size=(20,1), key='-NREGTO-'),
        sg.T('Статус', auto_size_text=True), sg.Combo(('Договор','Бронь','Подтверждено','Исполнена'), default_value=results[34], size=(12,1), key='-SREGTO-')],
        [sg.HorizontalSeparator()],
        [sg.T('Валюта тура', auto_size_text=True), sg.Combo(('Евро', 'Доллар', 'Рубль'), default_value=results[21], size=(10,1), key='-CURR-'), 
        sg.T('Стоимость (руб)', auto_size_text=True), sg.In(results[29],size=(7,1), key='-CTOURR-'),
        sg.T('Стоимость (вал)', auto_size_text=True), sg.In(results[28],size=(7,1), key='-CTOURC-'),
        sg.T('Аванс', auto_size_text=True), sg.In(results[30],size=(7,1), key='-VAVA-'), 
        sg.T('Курс оператора(Аванс)', auto_size_text=True), sg.In(results[31],size=(7,1), key='-CAVA-')],
        [sg.T('Дата аванса', auto_size_text=True), sg.In(results[22],size=(10,1), key='-DATEA-'),
#        sg.CalendarButton(button_text='', image_filename=path.join('ico', 'Calendar_24x24.png'), target='-DATEA-', format='%d.%m.%Y'),
        sg.Combo(('Получено','Оплачено','Нет'), default_value=results[23], size=(8,1), key='-SAVA-'),
        sg.T('Дата оплаты', auto_size_text=True), sg.In(results[24],size=(10,1), key='-DATEO-'),
#        sg.CalendarButton(button_text='', image_filename=path.join('ico', 'Calendar_24x24.png'), target='-DATEO-', format='%d.%m.%Y'),
        sg.Combo(('Получено','Оплачено','Нет'), default_value=results[25], size=(8,1), key='-SOPL-'),
        sg.T('Документы до', auto_size_text=True), sg.In(results[26], size=(10,1), key='-DATED-'),
#        sg.CalendarButton(button_text='', image_filename=path.join('ico', 'Calendar_24x24.png'), target='-DATED-', format='%d.%m.%Y'),
        sg.Combo(('Получены','Сданы','Выданы','Нет'), default_value=results[27], size=(8,1), key='-SDOC-')],
        [sg.HorizontalSeparator()],
        [sg.Frame('', frame_layout, element_justification = "center")]
#        [sg.Button('Договор') ,sg.Button('Сохранить'), sg.Button('Выход')]]
        ]
    rewnd = sg.Window('Информация по заявке', agencylayout, no_titlebar=False)

    while True:
        event, values =rewnd.read()
        if event == '-EXIT-'  or event is None:
            break

        if event == '-AHOTEL-': # добавление нового отеля в список по заявке
            answ = sg.popup('Добавить новый отель? ', custom_text=('Да', 'Нет'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Да':
#           Определение номера строки отеля (no_in_table)  
                cursor.execute('SELECT MAX(no_in_table) FROM req_accom WHERE id_req = "'+str(req_id)+'"')
                n_s_h = cursor.fetchone()
                if n_s_h[0] is None:
                    n_str_hotel = 1
                else:
                    n_str_hotel = n_s_h[0] + 1                
                ins_sql = "INSERT INTO req_accom (id_req, no_in_table, date_begin, date_end) VALUES ('" + str(req_id) + "','" + str(n_str_hotel)+ "','" +str(results[6]) + "','" +str(results[7]) + "');"
                cursor.execute(ins_sql)
                conn.commit()
                results3 = listhotel(cursor, req_id)
                reqhotelform(conn,results3[-1],req_id)
                results3 = listhotel(cursor, req_id)
                rewnd['-LHOTEL-'].update(values=results3)

        if event == '-MHOTEL-': # изменение данных по отелю
            if values['-LHOTEL-'] == []:
                nrow = 0
            else:
                nrow = values['-LHOTEL-'][0]
            reqhotelform(conn,results3[nrow],req_id)
            results3 = listhotel(cursor, req_id)
            rewnd['-LHOTEL-'].update(values=results3)

        if event == '-DHOTEL-': # удаление отеля из списка по заявке
            if values['-LHOTEL-'] == []:
                nrow = 0
            else:
                nrow = values['-LHOTEL-'][0]
            answ = sg.popup('Удалить отель ' + str(results3[nrow][3]) + ' из списка?', custom_text=('Да', 'Нет'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Да':
                del_sql = "DELETE FROM req_accom WHERE id_req = ? AND no_in_table = ?"
                column_values = (req_id, results3[nrow][0])
                cursor.execute(del_sql, column_values)
                conn.commit()
                results3 = listhotel(cursor, req_id)
                rewnd['-LHOTEL-'].update(values=results3)
        
        if event == '-DTURIST-':   # удаление туриста из заявки
            if values['-LTURISTS-'] == []:
                nrow = 0
            else:
                nrow = values['-LTURISTS-'][0]     
            answ = sg.popup('Удалить туриста ' + str(results4[nrow][1]) + ' из списка?', custom_text=('Да', 'Нет'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Да':
                del_sql = "DELETE FROM req_tourist WHERE id_req = ? AND id_cust = ?"
                column_values = (req_id, results4[nrow][0])
                cursor.execute(del_sql, column_values)
                conn.commit()
                results4 = listcust(cursor, req_id)
                rewnd['-LTURISTS-'].update(values=results4)

        if event == '-ATURIST-':   # удаление туриста из заявки
#            rewnd.Disable()
            tourist_id = custform.form(conn)
            if tourist_id != 0:
                ins_sql = "INSERT INTO req_tourist (id_req, id_cust) VALUES ('" + str(req_id) + "','" + str(tourist_id) + "');"
                cursor.execute(ins_sql)
                conn.commit()                
                results4 = listcust(cursor,req_id)
                rewnd['-LTURISTS-'].update(values=results4)      
#            rewnd.Enable()
#            rewnd.BringToFront()

        if event == '-CUST-':
#            rewnd.Disable()
            cust_id = custform.form(conn)
            if cust_id != 0:
                cursor.execute("SELECT fio FROM Cat_cust WHERE id_cust = "+ str(cust_id))
                results1 = cursor.fetchone()
                rewnd['-FIO-'].update(results1[0])
                answ = sg.popup('Добавить ' + results1[0] +' в список туристов по заявке?', custom_text=('Да', 'Нет'))
                if answ == 'Да':
                    ins_sql = "INSERT INTO req_tourist (id_req, id_cust) VALUES ('" + str(req_id) + "','" + str(cust_id) + "');"
                    cursor.execute(ins_sql)
                    conn.commit()
                    results4 = listcust(cursor,req_id)
                    rewnd['-LTURISTS-'].update(values=results4)
#            rewnd.Enable()
#            rewnd.BringToFront()

        if event == '-TOUROP-':
            try:
#                rewnd.Disable()
                id_oper = touropform.form(conn)
                if id_oper != 0:
                    cursor.execute("SELECT name_short_to FROM Cat_tourop WHERE id_to = "+ str(id_oper))
                    results2 = cursor.fetchone()
                    rewnd['-TO-'].update(results2[0])
#                rewnd.Enable()
#                rewnd.BringToFront()
            except:
                answ = sg.popup("ERROR", "-TOUROP-")

        if event == '-SAVE-':
            answ = sg.popup('Сохранить внесенные изменения? ', custom_text=('Сохранить', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Сохранить':
                if contformat.contdatef(values['-DATR-']) and contformat.contdatef(values['-DATEB-']) and contformat.contdatef(values['-DATEE-'])  and contformat.contdatef(values['-DATEA-'])  and contformat.contdatef(values['-DATEO-'])  and contformat.contdatef(values['-DATED-']):
                    column_values = (
                    values['-DATR-'], values['-NCONTR-'], values[0], values[1] , cust_id, values['-DATEB-'], values['-DATEE-'],\
                    values['-NNIGHT-'], values[2], values[3], values['-EPROG-'], values['-PRUS-'], values['-GID-'], values['-EXCUR-'],\
                    values['-LEAD-'], values['-VISA-'], values['-MEDS-'], values['-NSS-'], values['-NEVS-'], id_oper, values['-CURR-'],\
                    values['-DATEA-'], values['-SAVA-'], values['-DATEO-'], values['-SOPL-'], values['-DATED-'], values['-SDOC-'],\
                    values['-CTOURC-'], values['-CTOURR-'], values['-VAVA-'], values['-CAVA-'], values['-PRIM-'], values['-NREGTO-'],\
                    values['-SREGTO-'], req_id
                                )
                    upd_sql = "\
                    UPDATE list_request SET date_req = ?, numb_contr = ?, country = ?, region = ?, id_cust = ?, date_tour = ?, \
                    date_end_tour = ?, quant_night = ?, ticket = ?, transfer = ?, excur_prog  = ?, other_serv = ?, tour_guide = ?,\
                    transl_guide =  ?, team_leader = ?, visa = ?, med_ins = ?, acc_ins = ?, fail_ins = ?, id_to = ?, curr_tour = ?,\
                    date_prepay = ?, paid_prepay = ?, date_full_pay = ?, paid_full_pay = ?, date_doc = ?, rec_doc = ?, cost_tour_curr = ?,\
                    cost_tour_rub = ?, prepay_rub = ?, rate_to = ?, prim_ins = ?, numreq_tourop =?, status_req = ? WHERE id_req = ?\
                    "
                    cursor.execute(upd_sql,column_values)
                    conn.commit()
                else:
                    sg.popup('Ошибка в формате даты (правильный формат "дд.мм.гггг")')

        if event == '-DOGV-':
            answ = sg.popup('Сформировать договор по заявке? ', custom_text=('Сформировать', 'Отмена'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Сформировать':
                dogform(cursor, results, cust_id, id_oper, req_id)

        if event == '-ATRANS-': # добавление новой перевозки в список по заявке
            answ = sg.popup('Добавить новую перевозку? ', custom_text=('Да', 'Нет'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Да':
#           Определение номера строки транспорта (no_in_trans)  
                cursor.execute('SELECT MAX(no_in_trans) FROM req_trans WHERE id_req = "'+str(req_id)+'"')
                n_s_t = cursor.fetchone()
                if n_s_t[0] is None:
                    n_str_trans = 1
                else:
                    n_str_trans = n_s_t[0] + 1                
                ins_sql = "INSERT INTO req_trans (id_req, no_in_trans, date_there, date_back) VALUES ('" + str(req_id) + "','" + str(n_str_trans)+ "','" +str(results[6]) + "','" +str(results[7]) + "');"
                cursor.execute(ins_sql)
                conn.commit()
                results5 = listtrans(cursor, req_id)
                reqtransform(conn,results5[-1],req_id)
                results5 = listtrans(cursor, req_id)
                rewnd['-LTRANS-'].update(values=results5)

        if event == '-MTRANS-': # изменение данных по перевозке
            if values['-LTRANS-'] == []:
                nrow = 0
            else:
                nrow = values['-LTRANS-'][0]
            reqtransform(conn,results5[nrow],req_id)
            results5 = listtrans(cursor, req_id)
            rewnd['-LTRANS-'].update(values=results5)

        if event == '-DTRANS-': # удаление перевозки из списка по заявке
            if values['-LTRANS-'] == []:
                nrow = 0
            else:
                nrow = values['-LTRANS-'][0]
            answ = sg.popup('Удалить ' + str(results5[nrow][2]) + ' из списка?', custom_text=('Да', 'Нет'), button_type=sg.POPUP_BUTTONS_YES_NO)
            if answ == 'Да':
                del_sql = "DELETE FROM req_trans WHERE id_req = ? AND no_in_trans = ?"
                column_values = (req_id, results5[nrow][0])
                cursor.execute(del_sql, column_values)
                conn.commit()
                results5 = listtrans(cursor, req_id)
                rewnd['-LTRANS-'].update(values=results5)

    rewnd.close()
    return req_id
from datetime import datetime

def contdatef(date_text): # контроль введенной даты на формат
    corformat = True
    try:
        datetime.strptime(date_text, '%d.%m.%Y')
    except ValueError:
        corformat = False
    return corformat
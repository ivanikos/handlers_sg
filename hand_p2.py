import sys, io

buffer = io.StringIO()
sys.stdout = sys.stderr = buffer

import os
import eel
import csv

import weldlog_summary
import handler_beta
import handler_tracing
import update_dbs


@eel.expose
def read_path():
    with open(os.getcwd() + f'\\dbs\\path_to_summary.csv', 'r') as file:
        read_file = csv.reader(file, delimiter=",")
        for row in read_file:
            path_to_akt = row[0]

        res = f'<label class="form-label" >Введите путь куда сохранить сводку</label><input type="text" id="path_akt" placeholder="Введите путь куда сохранить сводку" value="{path_to_akt}">'

        return res

@eel.expose
def start_handler(path):
    print(path)
    try:
       handler_beta.start_handler(path)
       return "Сводка по ФАЗАМ 1, 2, 3 сформирована"
    except Exception as e:
        return f"Возникла ошибка! \n {e}"
@eel.expose
def start_handler_tracing(path):
    try:
        handler_tracing.create_summary_tracing(path)
        return "Сводка по по теплоспутникам сформирована"
    except:
        return "Возникла ошибка!"


@eel.expose
def start_handler_nkdkd(path):
    try:
        weldlog_summary.create_summary_nkdk(path)
        return "Сводка по % НК СГ сформирована"
    except:
        return "Возникла ошибка!"


@eel.expose
def update_bdtp():
    try:
        update_dbs.update_isotp_dbs()
        return "БД ТП по фазам обновлена успешно!"
    except:
        return "Возникла ошибка!"



if __name__ == '__main__':
    eel.init('front')
    eel.start('index.html', mode="chrome", size=(850, 700))
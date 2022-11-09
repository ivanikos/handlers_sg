import csv
import openpyxl as xl
import re
import xlsxwriter
import pandas as pd


# home laptop
# directory_dbs_files = r'C:\Users\vanik\PycharmProjects\handlers_sg\out_files_for_dbs\\'

# work laptop
directory_dbs_files = r'C:\Users\ignatenkoia\Desktop\work\GIT_PROJECTS\handlers_sg\dbs\\'



file_db_isotp = 'iso_tp_db.csv'


isotp_dic = {}
tp_dic = {}
iso_dic = {}

"""
Сводные списки для финальной записи в сводки по фазам.
"""
summary_iso_tp_phase_1 = [['Тестпакет', 'Изометрия', 'Линия', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Тип изоляции',
                           'Объём изоляции', 'RFI Мин.вата', 'RFI Металл. кожух', 'RFI Короб/чехол', 'RFI ДИГ',
                           'Статус уведомлений']]

summary_iso_tp_phase_2 = [['Тестпакет', 'Изометрия', 'Линия', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Тип изоляции',
                           'Объём изоляции', 'RFI Мин.вата', 'RFI Металл. кожух', 'RFI Короб/чехол', 'RFI ДИГ',
                           'Статус уведомлений']]

summary_iso_tp_phase_3 = [['Тестпакет', 'Изометрия', 'Линия', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Тип изоляции',
                           'Объём изоляции', 'RFI Мин.вата', 'RFI Металл. кожух', 'RFI Короб/чехол', 'RFI ДИГ',
                           'Статус уведомлений']]

summary_tp_phase_1 = [['Тестпакет', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Статус уведомлений']]

summary_tp_phase_2 = [['Тестпакет', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Статус уведомлений']]

summary_tp_phase_3 = [['Тестпакет', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Статус уведомлений']]


"""
Наполнение словарей по изометрии+тестпакет и по тестпакетам исходной подтверждённой информацией.
"""
with open(directory_dbs_files + file_db_isotp, 'r') as read_db:
    readed_db = csv.reader(read_db, delimiter=';')

    for row in readed_db:
        isotp_dic[row[0]] = ['', '', '', '', '', '', '', 0, '', '', '', '', '', '', '', '', '', '']
        tp_dic[row[2]] = ['', '', '', '', '', 0, '', '', '', '', '']
        iso_dic[row[1]] = ['', '', '', '']

with open(directory_dbs_files + file_db_isotp, 'r') as read_db:
    readed_db = csv.reader(read_db, delimiter=';')

    for row in readed_db:

        iso_with_tp = row[0].strip()
        isometric = row[1].strip()
        testpackage = row[2].strip()
        phase = row[3].strip()
        line = row[4].strip()
        title = row[5].strip()
        unit = row[6].strip()
        fluid = row[7].strip()
        ggn_status = row[8].strip()
        iso_length = float(row[9].strip().replace(',', '.'))
        rfi_erection = row[10].strip()
        rfi_test = row[11].strip()
        rfi_airblowing = row[12].strip()
        rfi_reinstatement = row[13].strip()
        type_ins = row[14].strip()
        volume_ins = row[15].strip()
        rfi_ins_cotton = row[16].strip()
        rfi_ins_metall = row[17].strip()
        rfi_ins_box = row[18].strip()

        isotp_dic[iso_with_tp] = [testpackage, isometric, line, title, unit, fluid, ggn_status, iso_length,
                                  rfi_erection, rfi_test, rfi_airblowing, rfi_reinstatement, type_ins,
                                  rfi_ins_cotton, rfi_ins_metall, rfi_ins_box, '', '']

        tp_dic[testpackage][0] = testpackage
        tp_dic[testpackage][1] = title
        tp_dic[testpackage][2] = unit
        tp_dic[testpackage][3] = fluid
        tp_dic[testpackage][4] = ggn_status
        tp_dic[testpackage][5] += iso_length
        tp_dic[testpackage][6] = rfi_erection
        tp_dic[testpackage][7] = rfi_test
        tp_dic[testpackage][8] = rfi_airblowing
        tp_dic[testpackage][9] = rfi_reinstatement

# for key in tp_dic.keys():
#     print(key, tp_dic[key])



# Проверка Журнал заявок АИС Р2 ФАЗА2-------------------------------------
df = pd.read_excel('Журнал заявок общий.xlsx')
df = df.sort_values(by='Дата подачи / Date of submission', ascending=True)
df.to_excel('Журнал заявок общий.xlsx', index=0)

print('Журнал заявок отсортирован по дате отработки инспекции.')

wb_journal_rfi = xl.load_workbook('Журнал заявок общий.xlsx')
sheet_journal_rfi = wb_journal_rfi['Sheet1']

replace_pattern_1 = ['-HT', '-VT', '-PT']
replace_pattern_2 = ['(T.T. REINSTATEMENT)', '(T.T. AIR BLOWING)', '(AIR BLOWING)', '(T.T AIR BLOWING',
                     '(T.T. ERECTION)', '(T.T .ERECTION)', '(T.T.TEST)', '(T.T. AIR BLOWIHG)', '(T.T. TEST)',
                     '(T.T ERECTION)', '(T.T TEST)', '(T.T REINSTATEMENT)',
                     '(T.T AIR BLOWING)', '(GPA AIR BLOWING)', '(GPA TEST)',
                     '(GPA ERECTION)', '(GPA REINSTATEMENT)', '(T.T. REISTATEMENT)', '(T.T.REINSTATEMENT)',
                     '(T.T RE-INSTATEMENT)', '( T.T AIR BLOWING )', '( T.T AIR BLOWING )',
                     '(T.T.ERECTION)', '(T.T.TEST)', '(T.T.AIR BLOWING)', '(T.T.REINSTATEMENT)']

res_summary = {} # ?????????

for i in sheet_journal_rfi['B2':'AO550000']:
    if i[0].value:
        rfi_number = str(i[1].value)
        tp_number = str(i[2].value)
        pkk = str(i[4].value)

        description_rfi = str(i[16].value)
        violation = str(i[35].value)
        name_insp = str(i[26].value)
        list_iso = str(i[8].value).split(';')
        volume_m = re.sub(r'[^0-9.]', '', str(i[18].value))
        category_cancelled = str(i[31].value)
        comment = str(i[39].value)

        for typo in replace_pattern_2:
            if typo in tp_number:
                tp_number = tp_number.replace(typo, '').strip()

        for typo in replace_pattern_1:
            if typo in tp_number:
                tp_number = tp_number.replace(typo, '').strip()

        print(tp_number)

        if 'Монтаж технологического трубопровода в рамках' in description_rfi:
            if 'Принято' in category_cancelled:

            pass


        if 'испытаний на прочность и плотность' in description_rfi:
            pass
        if 'испытаний технологического трубопровода  на прочность' in description_rfi:
            pass
        if 'испыт' and 'рочност' in description_rfi:
            pass
        if 'испытаний технологического трубопровода на прочность' in description_rfi:
            pass



        if 'родувка' in description_rfi and 'еплоспутн' not in description_rfi:
            pass

        if 'сборки технологических трубопроводов в проект' in description_rfi:
            pass
        if 'сборки технологических трубопроводов в рамках' in description_rfi:
            pass

        if 'дополнительных испытаний' in description_rfi:
            pass
        if 'дополн' in description_rfi:
            if 'Принято' in category_cancelled:
                pass
            else:
                if 'выдерж' in comment:
                    pass










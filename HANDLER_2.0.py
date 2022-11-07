import csv


# home laptop
# directory_dbs_files = r'C:\Users\vanik\PycharmProjects\handlers_sg\out_files_for_dbs\\'

# work laptop
directory_dbs_files = r'C:\Users\ignatenkoia\Desktop\work\GIT_PROJECTS\handlers_sg\dbs\\'



file_db_isotp = 'iso_tp_db.csv'


isotp_dic = {}
tp_dic = {}

"""
Сводные списки для финальной записи в сводки по фазам.
"""
summary_iso_tp_phase_1 = [['Тестпакет', 'Изометрия', 'Линия', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Тип изоляции',
                           'Объём изоляции', 'RFI Мин.вата', 'RFI Металл. кожух', 'RFI Короб/чехол',
                           'Статус уведомлений']]

summary_iso_tp_phase_2 = [['Тестпакет', 'Изометрия', 'Линия', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Тип изоляции',
                           'Объём изоляции', 'RFI Мин.вата', 'RFI Металл. кожух', 'RFI Короб/чехол',
                           'Статус уведомлений']]

summary_iso_tp_phase_3 = [['Тестпакет', 'Изометрия', 'Линия', 'Титул', 'Установка', 'Среда', 'Статус ГГН', 'Длина',
                           'RFI ERECTION', 'RFI TEST', 'RFI AIRBLOWING', 'RFI REINSTATEMENT', 'Тип изоляции',
                           'Объём изоляции', 'RFI Мин.вата', 'RFI Металл. кожух', 'RFI Короб/чехол',
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
        isotp_dic[row[0]] = ['', '', '', '', '', '', '', 0, '', '', '', '', '', '', '', '', '']
        tp_dic[row[2]] = ['', '', '', '', '', 0, '', '', '', '', '', '']

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
                                  rfi_ins_cotton, rfi_ins_metall, rfi_ins_box, '']

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

for key in tp_dic.keys():
    print(key, tp_dic[key])
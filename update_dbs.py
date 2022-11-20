# -*- coding: cp1251 -*-
import openpyxl as xl
import csv



# home laptop
file_isotp_db = r'C:\Users\vanik\PycharmProjects\handlers_sg\out_files_for_dbs\\'
out_db_isotp_dir = r'C:\Users\vanik\PycharmProjects\handlers_sg\dbs\\'

# work laptop
# out_db_isotp_dir = r'C:\Users\ignatenkoia\Desktop\work\GIT_PROJECTS\handlers_sg\dbs\\'
# file_isotp_db = r'C:\Users\ignatenkoia\Desktop\work\GIT_PROJECTS\handlers_sg\out_files_for_dbs'


# Создание файла .csv для дальнейшего использования-------------------------------------------
isotp_db = []

"""
Обновление БД ТП по всем фазам.
"""
def update_isotp_dbs():
    db_iso_tp = xl.load_workbook(file_isotp_db + "\\db_tp_p1,2,3_v2.0.xlsx")
    sheet_isotp = db_iso_tp['iso_tp_db']

    for i in sheet_isotp['A5':'W1000000']:
        if i[0].value:
            iso_with_tp = str(i[0].value).strip()
            isometric = str(i[1].value).strip()
            testpackage = str(i[2].value).strip()
            phase = str(i[3].value).strip()
            line = str(i[4].value).strip()
            title = str(i[5].value).strip()
            unit = str(i[6].value).strip()
            fluid = str(i[9].value).strip()
            ggn_status = str(i[13].value).strip()
            type_insulation = str(i[14].value).strip()
            iso_length = str(i[12].value).strip().replace(',', '.')

            if i[15].value:
                volume_insulation = str(i[15].value).strip()
            else:
                volume_insulation = 'n/d'

            rfi_erection = ''
            rfi_test = ''
            rfi_airblowing = ''
            rfi_reinstatement = ''

            rfi_insulation_mv = ''
            rfi_insulation_metall = ''
            rfi_insulation_box = ''

            if i[18].value:
                rfi_erection = str(i[18].value).strip()
            if i[19].value:
                rfi_test = str(i[19].value).strip()
            if i[20].value:
                rfi_airblowing = str(i[20].value).strip()
            if i[21].value:
                rfi_reinstatement = str(i[21].value).strip()

            if i[16].value:
                rfi_insulation_mv = str(i[16].value).strip()
            if i[17].value:
                rfi_insulation_metall = str(i[17].value).strip()
            if i[22].value:
                rfi_insulation_box = str(i[22].value).strip()

            isotp_db.append([iso_with_tp, isometric, testpackage, phase, line, title, unit, fluid, ggn_status,
                             iso_length, rfi_erection, rfi_test, rfi_airblowing, rfi_reinstatement,
                             type_insulation, volume_insulation, rfi_insulation_mv, rfi_insulation_metall, rfi_insulation_box])

    with open(out_db_isotp_dir + "iso_tp_db.csv", 'w', newline='') as writing_file:
            write_file = csv.writer(writing_file, delimiter=";")
            write_file.writerows(isotp_db)

    return print('БД ТП по фазам 1, 2, 3 успешно обновлена.')



update_isotp_dbs()
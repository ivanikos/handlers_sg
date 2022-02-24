import openpyxl as xl
import xlsxwriter
import datetime
import pandas as pd
import warnings
#warnings.simplefilter('ignore')
print('start')


result_WL = [['Тестпак', 'Линия', 'Изометрия', 'Стык', 'Процент контроля', 'Номер закл. УЗК ПО', 'Результат ПО',
              'Номер закл. РК ПО', 'Результат ПО', 'Дубль НК СГ']]

sorting_journal = pd.read_excel('Журнал P2 .xlsx')
sorting_journal = sorting_journal.sort_values(by='Дата РК', ascending=True)
sorting_journal.to_excel('Журнал P2 .xlsx')


wb_sg = xl.load_workbook('Журнал P2 .xlsx')
sheet_sg = wb_sg['Sheet1']
sg_rt_dic = {}
for k in sheet_sg['H2':'X240000']:
    if k[0].value:
        iso_number_sg = str(k[0].value).strip()
        joint_number_sg = str(k[5].value).strip()
        joint_status_rw = str(k[6].value).strip()
        xr_res_sg = str(k[16].value)
        sg_rt_dic[iso_number_sg+joint_number_sg+joint_status_rw] = [iso_number_sg, joint_number_sg, xr_res_sg]
    else:
        break
print('сг RT закончил')
wb_sg.close()
#----------------------------------------------------
#Журнал УЗК------------------------------------------
try:
    sorting_journal = pd.read_excel('Журнал УЗК труба.xlsx', sheet_name='P2')
    sorting_journal = sorting_journal.sort_values('Дата УЗК', ascending=True)
    sorting_journal.to_excel('Журнал УЗК труба.xlsx')
except:
    pass


wb_ut_sg = xl.load_workbook('Журнал УЗК труба.xlsx')
sheet_ut_sg = wb_ut_sg['Sheet1']

sg_ut_dic = {}
for i in sheet_ut_sg['E3':'P300000']:
    if i[0].value:
        iso_number_ut = str(i[0].value).strip()
        joint_number_ut = str(i[4].value).strip()
        ut_res_sg = str(i[11].value).strip()
        sg_ut_dic[iso_number_ut+joint_number_ut] = [iso_number_ut, joint_number_ut, ut_res_sg]
    else:
        break

print('сг UT закончил')
wb_ut_sg.close()

# Скан WeldLog'а -----------------

wb_WL = xl.load_workbook('WeldLog.xlsx')
sheet = wb_WL['WELDLOG']
wl_dic = {}
for i in sheet['D7':'CI200000']:
    if i[0].value:
        line_number = str(i[0].value).strip()
        iso_number = str(i[3].value).strip()
        control_percent = str(i[5].value)
        joint_number = str(i[9].value).strip() + str(i[10].value).strip()
        ut_number = str(i[27].value)
        ut_res = str(i[28].value)
        xr_number = str(i[30].value)
        xr_res = str(i[31].value)
        tp_short_code = str(i[83].value).strip()

        wl_dic[iso_number+joint_number] = [tp_short_code, iso_number, joint_number, control_percent, ut_number, ut_res,
                                           xr_number, xr_res, '']
    else:
        break

print('WL закончил')

print(len(wl_dic.keys()))
wl_summary_list = [['Тестпак', 'Isometric', 'Joint', 'Control %', 'Номер заключения УЗК ПО', 'Результат УЗК ПО',
                    'Номер заключения РК ПО', 'Результат РК ПО', 'Дубль контроль СГ']]

for key in wl_dic:
    double_nk_res = 'не проводился'
    if key in sg_ut_dic.keys():
        double_nk_res = sg_ut_dic[key][2]
    if key in sg_rt_dic.keys():
        double_nk_res = sg_rt_dic[key][2]

    wl_summary_list.append([wl_dic[key][0], wl_dic[key][1], wl_dic[key][2], wl_dic[key][3], wl_dic[key][4],
                            wl_dic[key][5], wl_dic[key][6], wl_dic[key][7], double_nk_res])




#Запись в файл выходных данных
workbook_wl = xlsxwriter.Workbook(f'WL_Phase2_DC_summary {datetime.datetime.now().strftime("%d.%m.%Y")}.xlsx')
ws = workbook_wl.add_worksheet()

ws.set_column(0, 0, 30)
ws.set_column(1, 1, 33)
ws.set_column(2, 2, 10)
ws.set_column(3, 3, 12)
ws.set_column(4, 4, 27)
ws.set_column(5, 5, 10)
ws.set_column(6, 6, 27)
ws.set_column(7, 7, 10)
ws.set_column(8, 8, 20)
ws.autofilter(f'A1:I{len(wl_dic.keys())}')

cell_form = workbook_wl.add_format()
cell_form.set_text_wrap(text_wrap=1)
for i, (short_tp, number_line, number_iso, number_joint, percent_control, number_ut, res_ut,
        number_xr, res_xr) in enumerate(wl_summary_list, start=1):
    color = cell_form

    ws.write(f'A{i}', short_tp, color)
    ws.write(f'B{i}', number_line, color)
    ws.write(f'C{i}', number_iso, color)
    ws.write(f'D{i}', number_joint, color)
    ws.write(f'E{i}', percent_control, color)
    ws.write(f'F{i}', number_ut, color)
    ws.write(f'G{i}', res_ut, color)
    ws.write(f'H{i}', number_xr, color)
    ws.write(f'I{i}', res_xr, color)



workbook_wl.close()
print('done')





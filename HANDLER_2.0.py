import csv


# home laptop
directory_dbs_files = r'C:\Users\vanik\PycharmProjects\handlers_sg\out_files_for_dbs\\'



file_db_isotp = 'iso_tp_db.csv'


isotp_dic = {}
with open(directory_dbs_files + file_db_isotp, 'r') as read_db:
    readed_db = csv.reader(read_db, delimiter=';')

    for row in readed_db:
        print(row)
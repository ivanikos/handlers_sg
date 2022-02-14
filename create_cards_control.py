# -*- coding: UTF-8 -*-

import os
import img2pdf

dir_name = r"C:\Users\ignatenkoia\Downloads\ИК ОК"

test = os.listdir(dir_name)




for root, dirs, files in os.walk(dir_name):
    for file in files:
        print(dir_name + "\\" + file)
        os.startfile(dir_name + "\\" + file)

        card_type = int(input())

        if card_type == 1:
            new_name = str(input())

            name_vik = 'ИК ВИК CPECC-CC-' + new_name + '.pdf'
            name_uzt = 'ИК УЗТ CPECC-CC-' + new_name + '.pdf'

            adress_file = (dir_name + "\\" + file)

            a4_page_size = [img2pdf.in_to_pt(8.3), img2pdf.in_to_pt(11.7)]
            layout_func = img2pdf.get_layout_fun(a4_page_size)

            pdf = img2pdf.convert(adress_file, layout_fun=layout_func)

            with open(r'C:\Users\ignatenkoia\Desktop\Акты ВИК Игнатенко\\' + name_vik, 'wb') as f:
                f.write(pdf)
            with open(r'C:\Users\ignatenkoia\Desktop\Акты ВИК Игнатенко\\' + name_uzt, 'wb') as f_1:
                f_1.write(pdf)

        else:
            new_name = str(input())

            name_vik = 'OK - ' + new_name + '.pdf'
            adress_file = (dir_name + "\\" + file)

            a4_page_size = [img2pdf.in_to_pt(8.3), img2pdf.in_to_pt(11.7)]
            layout_func = img2pdf.get_layout_fun(a4_page_size)

            pdf = img2pdf.convert(adress_file, layout_fun=layout_func)

            with open(r'C:\Users\ignatenkoia\Desktop\Акты ВИК Игнатенко\\' + name_vik, 'wb') as f:
                f.write(pdf)


#os.startfile(r"C:\Users\ignatenkoia\Downloads\Новая папка\WhatsApp Image 2021-10-28 at 11.54.47.jpeg")

"""names = os.listdir(os.getcwd())
for name in names:
    fullname = os.path.join(os.getcwd(), name)
    if os.path.isfile(fullname):
        print(fullname)"""
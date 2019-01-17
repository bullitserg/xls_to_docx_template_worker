import xlrd
from os.path import join, normpath, basename
from os import listdir, mkdir
from datetime import datetime
from docxtpl import DocxTemplate
from time import sleep
from shutil import move
from config import *

input_dir = normpath(input_dir)


def write_log(file, text):
    with open(file, mode='a', encoding='utf8') as log:
        log.write(str(text) + '\n')


def drop_one(number):
    return number - 1


while True:
    input_files = [file for file in listdir(input_dir) if (file.endswith('.xlsx') or file.endswith('.xls'))]
    if input_files:
        start_datetime = datetime.now()

        exec_out_dir = join(out_dir, start_datetime.strftime('%Y%m%d_%H%M%S'))
        mkdir(exec_out_dir)

        log_file = join(exec_out_dir, log_name)
        event = 'Starting %s' % start_datetime
        print(event)
        write_log(log_file, event)

        for file in input_files:
            file = join(input_dir, file)
            event = 'Working %s' % file
            write_log(log_file, event)

            excel_file = xlrd.open_workbook(file)
            excel_sheet = excel_file.sheet_by_index(0)
            parsed_data = list(excel_sheet.get_rows())[1:]

            for s in parsed_data:
                s = [v.value for v in s]

                money = s[drop_one(money_column)]
                if s[drop_one(date_column)] < '2019-01-01 00:00:00':
                    nds_percent = 18
                    nds_money = round(money * 0.18, 2)
                else:
                    nds_percent = 20
                    nds_money = round(money * 0.20, 2)

                procedures_d = {'procedure_number': s[drop_one(procedure_number_column)],
                                'name': s[drop_one(name_column)],
                                'inn': str(s[drop_one(inn_column)]).replace('.0', ''),
                                'address': s[drop_one(address_column)],
                                'money': str('{0:.2f}'.format(money).replace('.', ',')),
                                'nds_percent': nds_percent,
                                'nds_money': str('{0:.2f}'.format(nds_money).replace('.', ','))
                                }

                doc = DocxTemplate(template_file)
                doc.render(procedures_d)

                file_name = '_'.join([procedures_d['procedure_number'], procedures_d['inn'], procedures_d['name'].
                                     replace("\"", '').replace(' ', '_')[0:name_max_len] + '.docx'])

                doc.save(join(exec_out_dir, file_name))
                event = 'Create %s' % file_name
                write_log(log_file, event)

            move(file, exec_out_dir)

            event = 'File %s is done' % basename(file)
            print(event)
            write_log(log_file, event)

        end_datetime = datetime.now()
        event = 'Finished %s' % end_datetime
        print(event)
        write_log(log_file, event)

    sleep(awaiting_time)



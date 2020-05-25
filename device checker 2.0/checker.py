import xlrd
from colored import fg, bg, attr
from datetime import datetime

argus_path = 'data/argus.xlsx'
oracle_path = 'data/oracle.xlsx'
checked_path = 'data/checked.csv'

amount_set = dict()
dataset = dict()
trash_set = dict()
#exist_set = dict()

argus_name = 4
argus_code = 5
argus_sn = 8

oracle_name = 2
oracle_code = 1
oracle_amount = 25

argus_file = xlrd.open_workbook(argus_path)
argus_sheet = argus_file.sheet_by_index(0)

oracle_file = xlrd.open_workbook(oracle_path)
oracle_sheet = oracle_file.sheet_by_index(0)

for row_num in range(oracle_sheet.nrows):
    if oracle_sheet.cell_value(row_num, oracle_code) in amount_set:
        print(oracle_sheet.cell_value(row_num, oracle_code), 'Уже есть в словаре', amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Количество по ораклу'])
        print(amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Количество по ораклу'], oracle_sheet.cell_value(row_num, oracle_amount))
        amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Количество по ораклу'] += 5
        amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Количество по ораклу']
        print(oracle_sheet.cell_value(row_num, oracle_amount), 'текущее значение')
    else:
        amount_set.update({oracle_sheet.cell_value(row_num, oracle_code): {
            'Название по Ораклу': oracle_sheet.cell_value(row_num, oracle_name),
            'Количество по ораклу': oracle_sheet.cell_value(row_num, oracle_amount),
        }})

for row_num in range(argus_sheet.nrows):    # Создаём сводный словарь оракл + аргус
    if argus_sheet.cell_value(row_num, argus_code) in amount_set:
        oracle_data = amount_set[argus_sheet.cell_value(row_num, argus_code)]
    else:
        oracle_data = 'Данная позиция отсутствует в Оракле'

    dataset.update({argus_sheet.cell_value(row_num, argus_sn).lower(): {
        'Название по аргусу': argus_sheet.cell_value(row_num, argus_name),
        'Код по аргусу': argus_sheet.cell_value(row_num, argus_code),
        'Данные по ораклу': oracle_data
    }})

# вычитаем из словаря ракла готовые к списанию устройства.
try:
    exist_file = open(checked_path, 'r')
    print('%s\n\n\nПодгружены данные предыдущего отчёта%s' % (fg(6), attr(0)))
    for line in exist_file:
        device = line.strip().split(';')
        if device[2] in amount_set:
            amount_set[device[2]]['Количество по ораклу'] -= 1
        trash_set.update({device[0]: {device[1]: device[2]}})
        print(trash_set)
except FileNotFoundError:
    print('%s\n\n\nРанее созданные отчёты отсутствуют!%s' % (fg(4), attr(0)))


def add_line(line):
    try:
        checked_file = open(checked_path, 'a')
    except FileNotFoundError:
        checked_file = open(checked_path, 'a')
    checked_file.write(line+'\n')
    checked_file.close()
    return print('%sЗапись добавлена.%s' % (fg(28), attr(0)))


def serial_check(serial):

    if serial in dataset.keys():
        if dataset[serial]['Данные по ораклу'] == 'Данная позиция отсутствует в Оракле':
            return print('%sДанная позиция отсутствует в Оракле. Требуется корректировка!%s' % (fg(1), attr(0)))
        else:
            amount = dataset[serial]['Данные по ораклу']['Количество по ораклу']

        or_name = dataset[serial]['Данные по ораклу']['Название по Ораклу']
        ar_name = dataset[serial]['Название по аргусу']
        sn_code = dataset[serial]['Код по аргусу']
        amount_av = ''
        print(f'%sСерийрый номер - %s{serial} %s- найден!%s' %
              (fg(2), fg(3), fg(2), attr(0)))
        print(
            f'%sУстройство по аргусу - {ar_name}\nУстройство по ораклу - {or_name} %s' % (fg(2), attr(0)))
        #print(f'%sУстройство по аргусу- {ar_name}\nУстройство по ораклу {or_name} %s' % (fg(2), attr(0)))
        if ar_name != or_name:
            print('%sНе совпадает название.%s' % (fg(1), attr(0)))

        else:
            print(f'%sТекущее поличество по ораклу:%s {amount_set[sn_code]["Количество по ораклу"]}' % (
                fg(5), attr(0)))
            print(f'%sГотовим к списанию: %s{serial.upper()}%s' % (
                fg(3), fg(3), attr(0)))

            if amount_set[sn_code]['Количество по ораклу'] > 0:
                if serial in trash_set:
                    print('%sДанный серийный номер уже находится в списке к списанию!!!!!%s' % (
                        fg(1), attr(0)))
                else:
                    listok = [serial, dataset[serial]['Название по аргусу'],
                              dataset[serial]['Код по аргусу']]
                    line = ';'.join(listok)
                    add_line(line)
                    trash_set.update(
                        {serial: {dataset[serial]['Название по аргусу']: dataset[serial]['Код по аргусу']}})
                    amount_set[sn_code]['Количество по ораклу'] -= 1
                    print('Осталось к списанию:',
                          amount_set[sn_code]['Количество по ораклу'])

            elif amount < 1:
                print(f'%sНет едениц для списания.%s' % (fg(1), attr(0)))

    else:
        print(f'%sСерийный номер %s{serial} %s%s - не найден!%s' %
              (fg(1), fg(3), attr(0), fg(196), attr(0)))


print('%s%sДля завершения программы введите end \n\n\n%s' %
      (fg(2), attr(1), attr(0)))
current = input('Введите серийный номер: ').lower()

while current.lower() != 'end':
    # curent_check =
    serial_check(current)

    current = input('\n\n\nВведите серийный номер: ').lower()

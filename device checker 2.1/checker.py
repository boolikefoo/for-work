import xlrd
from colored import fg, bg, attr
from datetime import datetime

#таблицы с остатками
argus_path = 'data/argus.xls'
oracle_path = 'data/oracle.xls'
checked_path = 'data/checked.csv'

amount_set = dict()
dataset = dict()
trash_set = dict()
#exist_set = dict()

#Номер колонки по остаткам
argus_name = 4  #Название по аргусу
argus_code = 5  #Номенклатурный код по аргусу
argus_sn = 8    #Серийный номер по аргусу

oracle_name = 2 #Название по ораклу
oracle_code = 1 #Номенклатурый код по ораклу
oracle_amount = 27 #Количество остатков по ораклу (26 - колонка АА)
oracle_warehouse = 6    #Название склада по ораклу (складская организация)
oracle_ware_name = 7    #Складское подразделение
oracle_ware_orgname = 6 #Складская организация

#читаем таблицу аргуса
argus_file = xlrd.open_workbook(argus_path)
argus_sheet = argus_file.sheet_by_index(0)

#читаем таблицу оракла
oracle_file = xlrd.open_workbook(oracle_path)
oracle_sheet = oracle_file.sheet_by_index(0)

#считаем остатки по складам

for row_num in range(oracle_sheet.nrows):
    if oracle_sheet.cell_value(row_num, oracle_code) in amount_set:
        amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Количество по ораклу'] += oracle_sheet.cell_value(row_num, oracle_amount)
        if oracle_sheet.cell_value(row_num, oracle_warehouse) in amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Склад/остаток']:
            amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Склад/остаток'][oracle_sheet.cell_value(row_num, oracle_warehouse)] += oracle_sheet.cell_value(row_num, oracle_amount)
        else:
            amount_set[oracle_sheet.cell_value(row_num, oracle_code)]['Склад/остаток'].update({oracle_sheet.cell_value(row_num, oracle_warehouse) : oracle_sheet.cell_value(row_num, oracle_amount),})
    else:
        amount_set.update({oracle_sheet.cell_value(row_num, oracle_code): {
            'Название по Ораклу': oracle_sheet.cell_value(row_num, oracle_name),
            'Количество по ораклу': oracle_sheet.cell_value(row_num, oracle_amount),
            'Склад/остаток' : {oracle_sheet.cell_value(row_num, oracle_warehouse) : oracle_sheet.cell_value(row_num, oracle_amount)},
            'Складское подразделение' : oracle_sheet.cell_value(row_num, oracle_ware_name),
            'Складская организация' : oracle_sheet.cell_value(row_num, oracle_ware_orgname),
        }})

#print(amount_set)
# Создаём сводный словарь оракл + аргус
for row_num in range(argus_sheet.nrows):    
    if argus_sheet.cell_value(row_num, argus_code) in amount_set:
        oracle_data = amount_set[argus_sheet.cell_value(row_num, argus_code)]
    else:
        oracle_data = 'Данная позиция отсутствует в Оракле'

    dataset.update({argus_sheet.cell_value(row_num, argus_sn).lower(): {
        'Название по аргусу': argus_sheet.cell_value(row_num, argus_name),
        'Код по аргусу': argus_sheet.cell_value(row_num, argus_code),
        'Данные по ораклу': oracle_data
    }})

# вычитаем из словаря оракла готовые к списанию устройства.
try:
    exist_file = open(checked_path, 'r')
    print('%s\n\n\nПодгружены данные предыдущего отчёта%s' % (fg(6), attr(0)))
    for line in exist_file:
        device = line.strip().split(';')
        if device[2] in amount_set:
            amount_set[device[2]]['Количество по ораклу'] -= 1
        trash_set.update({device[0]: {device[1]: device[2]}})
        # print(trash_set)
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

def remove_line():
    pass

def serial_check(serial):
    answer = ''

    if serial in dataset.keys():
        if dataset[serial]['Данные по ораклу'] == 'Данная позиция отсутствует в Оракле':
            return print('%sДанная позиция отсутствует в Оракле. Требуется корректировка!%s' % (fg(1), attr(0)))
        else:
            amount = dataset[serial]['Данные по ораклу']['Количество по ораклу']

        or_name = dataset[serial]['Данные по ораклу']['Название по Ораклу']
        ar_name = dataset[serial]['Название по аргусу']
        sn_code = dataset[serial]['Код по аргусу']
        or_ware_name = dataset[serial]['Данные по ораклу']['Складское подразделение']
        or_ware_orgname = dataset[serial]['Данные по ораклу']['Складская организация']

        print(f'%sСерийрый номер - %s{serial} %s- найден!%s' %
              (fg(2), fg(3), fg(2), attr(0)))
        print(dataset[serial]['Данные по ораклу']['Склад/остаток'])
        print(
            f'%sУстройство по аргусу - {ar_name}\nУстройство по ораклу - {or_name} %s' % (fg(2), attr(0)))
        #print(f'%sУстройство по аргусу- {ar_name}\nУстройство по ораклу {or_name} %s' % (fg(2), attr(0)))
        if ar_name != or_name:
            print('%sНе совпадает название.%s' % (fg(1), attr(0)))
            answer = input('Подготовить к списанию? Y/N: ').lower()

        print(f'%sТекущее поличество по ораклу:%s {amount_set[sn_code]["Количество по ораклу"]}' % (
            fg(5), attr(0)))
        print(f'%sГотовим к списанию: %s{serial.upper()}%s' % (
            fg(3), fg(3), attr(0)))

        if amount_set[sn_code]['Количество по ораклу'] > 0:
            if serial in trash_set:
                print('%sДанный серийный номер уже находится в списке к списанию!!!!!%s' % (
                    fg(1), attr(0)))
            elif answer == 'n':
                answer = ''
            else:

                listok = [serial, dataset[serial]['Название по аргусу'],
                          dataset[serial]['Код по аргусу'],
                          dataset[serial]['Данные по ораклу']['Складское подразделение'],
                          dataset[serial]['Данные по ораклу']['Складская организация'],
                          str(dataset[serial]['Данные по ораклу']['Склад/остаток'].items()),
                          ]
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

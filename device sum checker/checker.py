import xlrd
import datetime

#Before start, upload in derectory argus and oracle devices lists in excel sheets
#init vars and dicts

dev_exist = {}
new_devices = {}
dev_exist_names = {}
dev_balance = {}
balance_counter = {}
a = ''
find = False
s = str(datetime.date.today())

# Give the location of the file 
argus_file = ("argus.xlsx")     #devices list from argus
oracle_file = ("oracle.xlsx")   #devices list from oracle
  
# Open Workbook argus 
wba = xlrd.open_workbook(argus_file) 
sheet = wba.sheet_by_index(0) 

#open WorkBook oracle
wbo = xlrd.open_workbook(oracle_file)
sheet_or = wbo.sheet_by_index(0)


# For row 0 and column 0 
#sheet.cell_value(0, 0)

#make devices dict
for i in range(sheet.nrows):
    if sheet.cell_value(i,5) in dev_exist.keys():
        dev_exist[sheet.cell_value(i,5)].append(sheet.cell_value(i,8))
        

    else:
        dev_exist[sheet.cell_value(i,5)] = [sheet.cell_value(i,8)]

for i in range(sheet.nrows):
    dev_exist_names[sheet.cell_value(i,5)] = sheet.cell_value(i,4)

#make balance dict
for i in range(sheet_or.nrows):
    if sheet_or.cell_value(i,1) in dev_balance.keys():
        dev_balance.update({sheet_or.cell_value(i, 1) : dev_balance[sheet_or.cell_value(i, 1)] + sheet_or.cell_value(i,16)})
    else:
        dev_balance[sheet_or.cell_value(i, 1)] = sheet_or.cell_value(i, 16)

print("для завершения программы введите end")

while a != 'end':
    a = input("Введите серийный номер: ")
    for i in dev_exist.keys():
        for j in dev_exist[i]:
            #print(j)
            if a == j:
                print(f"Серийный номер {j} найден\nСоответствует коду -- {i} -- \nCоответствует наименованию -- {dev_exist_names[i]}\n", f"")
                find = True
                if i in new_devices.keys():
                    new_devices[i].append(a)
                    balance_counter.update({i:balance_counter[i]+1})

                else:               
                    new_devices[i] = [a]
                    balance_counter[i] = 1

                if i in dev_balance.keys():
                    if dev_balance[i] >= balance_counter[i]:                 
                        print(f"Можно списать данных устройств: {dev_balance[i]} , к списанию уже готово {balance_counter[i]}.")
                    else:
                        print(f"Вы привысили максимальное число устройств для списания по текущей позиции: {i}")
                        new_devices[i].pop()
                        balance_counter.update({i:balance_counter[i]-1})
                elif i not in dev_balance.keys():
                    print(f"Требуется корректировка номенклатурного кода устройства:{i}")

                break
    if find:
        
        find = False
    else:
        if a != 'end':
            print('Серийный номер не найден! Поищем снова?')

with open(s + " - checked.txt", "w") as outf:
    for key, val in new_devices.items():
        outf.write(f'{key}:{val}\n')
#print(dev_exist_names.items())    
#print(balance_counter)

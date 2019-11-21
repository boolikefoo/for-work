import xlrd
import datetime

dev_exist = {}
new_devices = {}
# Give the location of the file 
loc = ("dev_list.xlsx") 
  
# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 
sheet.cell_value(0, 0)

#with open("currentList.txt", 'r') as curentList:

for i in range(sheet.nrows):
    if sheet.cell_value(i,5) in dev_exist.keys():
        dev_exist[sheet.cell_value(i,5)].append(sheet.cell_value(i,8))
        

    else:
        dev_exist[sheet.cell_value(i,5)] = [sheet.cell_value(i,8)]


a = ''
find = False
print("для завершения программы введите end")
while a != 'end':
    a = input("Введите серийный номер: ")
    for i in dev_exist.keys():
        for j in dev_exist[i]:
            #print(j)
            if a == j:
                print("Серийный номер найден", j)
                find = True
                if i in new_devices.keys():
                    new_devices[i].append(a)
                else:               
                    new_devices[i] = [a]
                break
                print("breaknull")
            
    if find:
        
        find = False
    else:
        print('Серийный номер не найден! Поищем снова?')
s = str(datetime.date.today())
with open(s + " - checked.txt", "w") as outf:
    for key, val in new_devices.items():
        outf.write('{}:{}\n'.format(key,val))
#print(s)    

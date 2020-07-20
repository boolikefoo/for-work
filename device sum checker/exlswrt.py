import xlrd

dev_name = {}
dev_balance = {}
balance_counter = {}
dev_exist_names = {}
a = ''
find = False

# Give the location of the file 
oracle_file = ("oracle.xlsx")   #devices list from oracle

#open WorkBook oracle
wbo = xlrd.open_workbook(oracle_file)
sheet_or = wbo.sheet_by_index(0)

dct = {}

with open("2020-04-13 - checked.txt", "r") as inf:
    for line in inf:
        (key, val) = line.strip("\n").split(":")
        dct[key] = val.strip("[']").split("', '")

for i in range(sheet_or.nrows):
    '''
    if sheet_or.cell_value(i,1) in dev_exist_names.keys():
        dev_exist_names[sheet_or.cell_value(i,1)].append(sheet_or.cell_value(i,2))
        

    else:
        '''
    dev_exist_names[sheet_or.cell_value(i,1)] = sheet_or.cell_value(i,2)

for key in dct.keys():
    for val in dct[key]:
        #print(val)   #Печатаем SN 
        
        if key in dev_exist_names.keys(): #Получаем наименования оборудования
            print(dev_exist_names[key] )
        #print(key)   #Печатаем номенклатурный код   
        
#print(dev_exist_names)


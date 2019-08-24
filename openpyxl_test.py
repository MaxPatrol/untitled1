from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from openpyxl import *
from tkinter import filedialog
from tkinter import *



inotab = "ETOPAHKCBeopakc"
outtab = "ЕТОРАНКСВеоракс"
trantab = str.maketrans(inotab, outtab)

#str = "ghbdtn vbh"
#print (str.translate(trantab, 'xm'))


#Modules - cписок, содержит списки модулей в каждом шкафу
#SelectedFiles - список имен файлов, выбранных через диалоговое окно

root = Tk()

SelectedFiles = filedialog.askopenfilenames(initialdir = "/",title = "Выбор файлов серийников модулей",filetypes = (("xlsx","*.xlsx"),("all files","*.*")))
print ("Выбраны файлы:")
for i in SelectedFiles:
    print(i)
print ("")

Modules =[]
NamesBoxs = []
for file in SelectedFiles:
    print("Обрабатывается файл {0} ...".format(file))
    wb = load_workbook(file)
    ws = wb[wb.sheetnames[0]]

    print("Поиск имени шкафа ...")
    #Поиск имени в первой строке
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            temp = cell.value
            #print(temp)
            if temp != None:
                if re.search('[a-zA-Z]', temp):
                    print("АХТУНГ!! {0}, {1}".format(temp, re.findall('[a-zA-Z]', temp)))
                    print("АХТУНГ!!. Исправление...")
                    temp = temp.translate(trantab)
                    if re.search('[a-zA-Z]', temp):
                        print("АХТУНГ!!. Исправление...Провал!")
                    else:
                        print("АХТУНГ!!. Исправлено!")
                if "ПТ" in temp:
                    print("Найдено имя шкафа: {0}".format(cell.value))
    print("")

    print("Поиск модулей...")
    b = []
    typeModule=''
    for row in ws.rows:

        placeModule = row[1].value
        typeModule = row[2].value

        if placeModule != None and len(placeModule)>3 and typeModule != None and len(typeModule) > 2:
            #print("Обрабатываем строку")
            if placeModule[0] == 'A' and (placeModule[2] == '.' or placeModule[3] == '.') and typeModule[:2] != 'CH':
                b.append([placeModule,typeModule])
                print ([placeModule,typeModule])
    Modules.append(b)
    print("")

for case in Modules:
    for mod in case:
        print(mod)













wb = load_workbook(filename='ser ПТ-2.КЦ.xlsx')

ws = wb.active

for row in ws.values:
    for value in row:
        s = ''
        if ("ПТ" in str(value)):
            print('ups'+str(value)[str(value).find("ПТ"):str(value).find("ПТ")+3])
            #li.append(str(valuse))
        print(value)
    # print(row)




# Номер последней корзины
LastRack = 8
###########################
###########################
###########################


# Cоздаем файл для записи результатов
wb = Workbook()

dest_filename = 'empty_book.xlsx'
ws = wb.active

#Шаблон для КЦ(обязательный)

#Серая заливка неиспользуемых ячеек

a= PatternFill("solid", fgColor="808080")
for i in range(36):
    tc = ws.cell(row=i + 1, column=1)
    tc.fill = a
    tc = ws.cell(row=i + 1, column=31)
    tc.fill = a
    tc = ws.cell(row=i + 1, column=32)
    tc.fill = a
    tc = ws.cell(row=i + 1, column=33)
    tc.fill = a

for i in range(30):
    for i2 in range(5, 36,6):
        tc = ws.cell(row=i2, column=i+1)
        tc.fill = a
        tc = ws.cell(row=i2+1, column=i+1)
        tc.fill = a

# Установка ширины столбцов
ws.column_dimensions[get_column_letter(1)].width = 7
ws.column_dimensions[get_column_letter(2)].width = 14
for i in range(29):
    ws.column_dimensions[get_column_letter(3+i)].width = 8

# Установка высоты строк
for i in range(1, 36,6):
    ws.row_dimensions[i].height = 11
    ws.row_dimensions[i+1].height = 73
    ws.row_dimensions[i+2].height = 9
    ws.row_dimensions[i+3].height = 12
    ws.row_dimensions[i+4].height = 12
    ws.row_dimensions[i+5].height = 12

# Расставляем объедененные ячейки со стилями
thin = Side(border_style='medium', color="000000")
all_border_cell = Border(top=thin, left=thin, right=thin, bottom=thin)
centred_cell_style = Alignment(horizontal="center", vertical="center")

a = [['A1',1],['A2',7],['A3',13],['A4',19],['A5',25],['A6',31]]
for i in a:
    tc = ws.cell(i[1],3,i[0])
    tc.border = all_border_cell;
    tc.alignment = centred_cell_style
    ws.merge_cells(start_row=i[1], start_column=3, end_row=i[1]+3, end_column=3)

a = [['КЦ',1],['КС',13],['КК',25]]
for i in a:
    tc = ws.cell(i[1],30,i[0])
    tc.border = all_border_cell;
    tc.alignment = centred_cell_style
    ws.merge_cells(start_row=i[1], start_column=30, end_row=i[1]+9, end_column=30)




wb.save(filename=dest_filename)

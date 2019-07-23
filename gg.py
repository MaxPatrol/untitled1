from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.utils import get_column_letter

wb = load_workbook(filename='d:\\1.xlsx')

ws = wb.active

li = []
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


# создаем файл для записи результатов
wb = Workbook()

dest_filename = 'd:\\empty_book.xlsx'
ws = wb.active



# Определяем стили

thin = Side(border_style='medium', color="000000")
centred_cell_style = Alignment(horizontal="center", vertical="center")
all_border_cell = Border(top=thin, left=thin, right=thin, bottom=thin)
# Составляем шаблон для 'Компоновка корзин'


# Установка ширины столбцов
ws.column_dimensions[get_column_letter(1)].width = 7
ws.column_dimensions[get_column_letter(2)].width = 14
for i in range(29):
    ws.column_dimensions[get_column_letter(3+i)].width = 8

# Корзина 1 ---------------------------------------------------------------------
for i in range(6):
    tc = ws.cell(row=i+1, column=1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1, column=32)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1, column=33)
    tc.fill = PatternFill("solid", fgColor="808080")
for i in range(29):
    tc = ws.cell(row=5, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=6, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
# Установка высоты строк
ws.row_dimensions[1].height = 11
ws.row_dimensions[2].height = 73
ws.row_dimensions[3].height = 9
ws.row_dimensions[4].height = 12
ws.row_dimensions[5].height = 12
ws.row_dimensions[6].height = 12

# Объединения

ws['C1'] = 'A1'
ws['C1'].border = all_border_cell
ws['C1'].alignment = centred_cell_style
ws.merge_cells('C1:C4')

# Корзина 2 ---------------------------------------------------------------------
for i in range(6):
    tc = ws.cell(row=i+6, column=1)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i+6, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i + 1+6, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+6, column=32)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+6, column=33)
    tc.fill = PatternFill("solid", fgColor="808080")

for i in range(31):
    tc = ws.cell(row=11, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=12, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
ws.row_dimensions[7].height = 11
ws.row_dimensions[8].height = 73
ws.row_dimensions[9].height = 9
ws.row_dimensions[10].height = 12
ws.row_dimensions[11].height = 12
ws.row_dimensions[12].height = 12

# Объединения

ws['C7'] = 'A2'
ws['C7'].border = all_border_cell
ws['C7'].alignment = centred_cell_style
ws.merge_cells('C7:C10')

ws['AD1'] = 'КЦ'
ws['AD1'].border = all_border_cell
ws['AD1'].alignment = centred_cell_style
ws.merge_cells('AD1:AD10')

# Корзина 3 ---------------------------------------------------------------------
for i in range(6):
    tc = ws.cell(row=i+12, column=1)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i+12, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i + 1+12, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+12, column=32)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+12, column=33)
    tc.fill = PatternFill("solid", fgColor="808080")

for i in range(29):
    tc = ws.cell(row=17, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=18, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
ws.row_dimensions[13].height = 11
ws.row_dimensions[14].height = 73
ws.row_dimensions[15].height = 9
ws.row_dimensions[16].height = 12
ws.row_dimensions[17].height = 12
ws.row_dimensions[18].height = 12

# Объединения

ws['C13'] = 'A3'
ws['C13'].border = all_border_cell
ws['C13'].alignment = centred_cell_style
ws.merge_cells('C13:C16')

# Корзина 4 ---------------------------------------------------------------------
for i in range(6):
    tc = ws.cell(row=i+18, column=1)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i+18, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i + 1+18, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+18, column=32)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+18, column=33)
    tc.fill = PatternFill("solid", fgColor="808080")

for i in range(31):
    tc = ws.cell(row=23, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=24, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
ws.row_dimensions[19].height = 11
ws.row_dimensions[20].height = 73
ws.row_dimensions[21].height = 9
ws.row_dimensions[22].height = 12
ws.row_dimensions[23].height = 12
ws.row_dimensions[24].height = 12

# Объединения

ws['C19'] = 'A4'
ws['C19'].border = all_border_cell
ws['C19'].alignment = centred_cell_style
ws.merge_cells('C19:C22')


ws['AD13'] = 'КС'
ws['AD13'].border = all_border_cell
ws['AD13'].alignment = centred_cell_style
ws.merge_cells('AD13:AD22')

# Корзина 5 ---------------------------------------------------------------------
for i in range(6):
    tc = ws.cell(row=i+24, column=1)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i+24, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i + 1+24, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+24, column=32)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+24, column=33)
    tc.fill = PatternFill("solid", fgColor="808080")

for i in range(29):
    tc = ws.cell(row=29, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=30, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
ws.row_dimensions[25].height = 11
ws.row_dimensions[26].height = 73
ws.row_dimensions[27].height = 9
ws.row_dimensions[28].height = 12
ws.row_dimensions[29].height = 12
ws.row_dimensions[30].height = 12

# Объединения

ws['C25'] = 'A5'
ws['C25'].border = all_border_cell
ws['C25'].alignment = centred_cell_style
ws.merge_cells('C25:C28')

# Корзина 6 ---------------------------------------------------------------------
for i in range(6):
    tc = ws.cell(row=i+30, column=1)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i+30, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")

    tc = ws.cell(row=i + 1+30, column=31)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+30, column=32)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=i + 1+30, column=33)
    tc.fill = PatternFill("solid", fgColor="808080")

for i in range(31):
    tc = ws.cell(row=35, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
    tc = ws.cell(row=36, column=i+1)
    tc.fill = PatternFill("solid", fgColor="808080")
ws.row_dimensions[31].height = 11
ws.row_dimensions[32].height = 73
ws.row_dimensions[33].height = 9
ws.row_dimensions[34].height = 12
ws.row_dimensions[35].height = 12
ws.row_dimensions[36].height = 12

# Объединения

ws['C31'] = 'A6'
ws['C31'].border = all_border_cell
ws['C31'].alignment = centred_cell_style
ws.merge_cells('C31:C34')


ws['AD25'] = 'КK'
ws['AD25'].border = all_border_cell
ws['AD25'].alignment = centred_cell_style
ws.merge_cells('AD25:AD34')

wb.save(filename=dest_filename)

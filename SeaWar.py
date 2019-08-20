import time, threading
from tkinter import *
from tkinter import messagebox

x = 0
mouseX = 0
mouseY = 0
cell = ''
maps = []
A = ['A', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж', 'З', 'И', 'К']
D = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
##############################################################


def infinite_process():
    global x,cell
    x = x + 1
    print("Infinite Loop"+str(x))
    root.after(500, infinite_process)
    #c.itemconfig(maps[x], fill='green', width=2)
    if cell:
        a = A.index(cell[0])
        b = (D.index(int(cell[1])))
        print(a*10+b)
        c.itemconfig(maps[a*10+b], fill='green', width=2)
        cell = ''

##############################################################


def canvas_lmb(event):
    global mouseX, mouseY, cell, A, D

    mouseX = root.winfo_pointerx() - root.winfo_rootx() - 50
    mouseY = root.winfo_pointery() - root.winfo_rooty() - 50
    # print("Курсор находится на позиции х={} y={}".format(x, y))
    # print(x//50, y//50)
    cell = A[mouseX//50]+str(D[mouseY//50])
    # print(cell)

    # c.itemconfig(maps[(x//50)*10+(y//50)], fill='green', width=2)


def movem(event):
    global mouseX, mouseY, cell, A, D

    mouseX = root.winfo_pointerx() - root.winfo_rootx() - 50
    mouseY = root.winfo_pointery() - root.winfo_rooty() - 50
    # print("Курсор находится на позиции х={} y={}".format(x, y))
    # print(x//50, y//50)

    # print(cell)

    #c.itemconfig(l5, text=A[mouseX//50]+str(D[mouseY//50]))
    l5['text'] = A[mouseX // 50] + str(D[mouseY // 50])

root = Tk()

root.geometry('800x560')

c = Canvas(root, width=560, height=560)
c.pack(side=LEFT)
l4 = Label(root, width=25, height=4, bg='lightblue', text="Фаза 1")
l4.pack(side=TOP)
l4 = Label(root, width=25, height=4, bg='lightgreen', text="Игрок. Расстановка")
l4.pack(side=TOP)
l5 = Label(root, width=25, height=1, bg='lightgreen', text="Позиция")
l5.pack(side=BOTTOM)

b = Button(root, text="Старт")
b.pack()


for y in range(1, 11):
    for x in range(1, 11):
        a = c.create_rectangle(5+50*y, 5+50*x, 51+50*y, 51+50*x, fill='white', outline='black',
        width=2, activedash=(5, 3))
        maps.append(a)

for x in range(1, 11):
    a = c.create_text(26+50*x, 29, text=A[x-1], justify=CENTER, font="Verdana 24")
    a = c.create_text(26, 26+50*x, text=str(x), justify=CENTER, font="Verdana 24")


c.bind('<Button-1>', canvas_lmb)
c.bind('<Motion>', movem)
root.after(500, infinite_process)
root.mainloop()

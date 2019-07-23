from tkinter import *
from tkinter import messagebox as mb
from time import *


maps =[]

def change(event):

    x = root.winfo_pointerx() - root.winfo_rootx()
    y = root.winfo_pointery() - root.winfo_rooty()
    print("Курсор находится на позиции х={} y={}".format(x, y))
    print(x//50, y//50)

    c.itemconfig(maps[(x//50)*10+(y//50)], fill='green', width=2)

root = Tk()

c = Canvas(root, width=510, height=510)
c.pack()
b = Button(text="Старт")
b.pack()


for y in range(10):
    for x in range(10):
        a = c.create_rectangle(5+50*y, 5+50*x, 51+50*y, 51+50*x, fill='white', outline='black',
        width=2, activedash=(5, 3))
        maps.append(a)


mb.showerror("Ошибка", "Должно быть введено число")

c.bind('<Button-1>', change)


root.mainloop()


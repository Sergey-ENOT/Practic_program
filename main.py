from tkinter import *
                  #TODO Сделать 5 новшеств(обработка или что-то новое из функций)
                  #TODO using openpyxl для чтения файла
root = Tk()
root.config(background='#2f944a')
root.title("Welcome")
root.geometry('1200x760')
date_ot = Label(root, text='от')
date_start = Entry(root)
date_do = Label(root, text='до')
date_finish = Entry(root)
analysis = Button(root, text='Анализ', height=2, width=12, font=("Times", 11))
save_button = Button(root, text='Сохранить', height=2, width=12, font=("Times", 11))
head_program = Label(root, text='АНАЛИЗ УЯЗВИМОСТЕЙ', width=760, background='#cf2b1f', font=("Times", 20))

r_var = BooleanVar()
r_var.set(0)
button_time_set = Radiobutton(root, text='за определённый период', variable=r_var, value=0)
button_time_all = Radiobutton(root, text='за всё время', variable=r_var, value=1)

low_hole = Label(root, background='yellow', text='Низкий', width=12, font=("Times", 14))
medium_hole = Label(root, background='yellow', text='Средний', width=12, font=("Times", 14))
high_hole = Label(root, background='yellow', text='Высокий', width=12, font=("Times", 14))
crit_hole = Label(root, background='yellow', text='Критический', width=12, font=("Times", 14))


# def EntryToLabel(event):
#     s = date_start.get()
#     l['text'] = s


#b.bind('<Button-1>', EntryToLabel)

head_program.pack(padx=0, pady=0)
low_hole.place(x=20, y=300)
medium_hole.place(x=20, y=340)
high_hole.place(x=20, y=380)
crit_hole.place(x=20, y=420)
date_start.place(x=350, y=40)
date_finish.place(x=520, y=40)
date_ot.place(x=320, y=40)
date_do.place(x=490, y=40)
analysis.place(x=900, y=285)
save_button.place(x=900, y=340)

button_time_set.pack(anchor=W)
button_time_all.pack(anchor=W)


root.mainloop()

from tkinter import *
import docx
import xlrd2
from tkinter import messagebox
from datetime import datetime
from docx import Document

root = Tk() # Главное окно
root.title("Welcome") # Название окна
root.geometry("600x350") # Размер окна
root.configure(background='#f5f5f5') # Цвет заднего фона окна
# Создание интерфейса
radioButtonDateVar = BooleanVar() # Создание радиокнопок
radioButtonDateVar.set(0)
radioButtonDateOn = Radiobutton(text="По дате", bg='#FFFAFA', variable=radioButtonDateVar, value=1)
radioButtonDateOff = Radiobutton(text="За все время", bg = '#FFFAFA', variable=radioButtonDateVar, value=0)
# Создание кнопок, полей, лейблов
buttonAnalysis = Button(root, bg='#008B8B', font='Times 12', text="Анализ", width=13, height=2)
buttonClear = Button(root, bg='#008B8B', font='Times 12', text="Удалить", width=13, height=2)
buttonSave = Button(root, bg='#008B8B', font='Times 12', text="Сохранить", width=13, height=2)
labelLow = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Низкий")
labelLowOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelMid = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Средний")
labelMidOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelHigh = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Высокий")
labelHighOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelSuper = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Критический")
labelSuperOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelDate = Label(root, text="Введите необходимую дату:", state=DISABLED,
                  bg='#FFFAFA',font='Times 13', fg='#000', width=30)
labelFromDate = Label(root, text=" От:", state=DISABLED, bg='#FFFAFA', fg='black', width=5)
labelToDate = Label(root, text="До:", state=DISABLED, bg='#FFFAFA', fg='black', width=5)
labelDateInfo = Label(root, text="Анализ уязвимостей WORD", bg='#008080', font='Times 20', fg='#999', width=50)
labelToInfo = Label(root, bg='#FFFAFA', fg='black', width=20)
textBoxFromDate = Entry(root, state=DISABLED, width=10)
textBoxToDate = Entry(root, state=DISABLED, width=10)


def dateOn(event): # Функция для радиокнопки "По дате", включает поля для ввода даты.
    labelDateInfo.configure(state=NORMAL)
    textBoxFromDate.configure(state=NORMAL)
    textBoxToDate.configure(state=NORMAL)
    labelFromDate.configure(state=NORMAL)
    labelToDate.configure(state=NORMAL)


def dateOff(event): # Функция для радиокнопки "Без даты", выключает и очищает поля для ввода даты.
    textBoxFromDate.delete(0, END)
    textBoxToDate.delete(0, END)
    labelDateInfo.configure(state=DISABLED)
    textBoxFromDate.configure(state=DISABLED)
    textBoxToDate.configure(state=DISABLED)
    labelFromDate.configure(state=DISABLED)
    labelToDate.configure(state=DISABLED)


def AnalysisWithDate(event): # Функция для проверки правильности ввода даты
    radio_condition = radioButtonDateVar.get() # Заносим в переменную radio_condition состояние радиокнопок (1 или 0)

    if radio_condition == 1: # Если радиокнопка "По дате" включена (1)
        dataFrom = textBoxFromDate.get()
        dataTo = textBoxToDate.get()
        #dataFrom = '12.12.2010'
        #dataTo = '12.12.2015'
        if len(dataFrom and dataTo) == 10 and (dataFrom[2] and dataTo[2]) == '.' and \
                (dataFrom[5] and dataTo[5]) == '.' and dataFrom[6:].isnumeric() and dataTo[6:].isnumeric() and \
                dataFrom[:2].isnumeric() and dataTo[:2].isnumeric() and dataFrom[3:5].isnumeric() and \
                dataTo[3:5].isnumeric():
            tsFrom = datetime(year=int(dataFrom[6:]), month=int(dataFrom[3:5]), day=int(dataFrom[:2]))
            tsTo = datetime(year=int(dataFrom[6:]), month=int(dataFrom[3:5]), day=int(dataFrom[:2]))
            if (tsFrom.date and tsTo.day) > 0 and (tsFrom.day and tsTo.day) < 32 and \
                (tsFrom.month and tsTo.month) > 0 and \
                (tsFrom.month and tsTo.month) < 13 and \
                (tsFrom.year and tsTo.year) > 1900:
                Analysis(event)
            else:
                messagebox.showerror('Ошибка', 'Некорректно введена дата') # Если дата введена некорректно - выводим окно с ошибкой
        else:
            messagebox.showerror('Ошибка', 'Некорректно введена дата') # Если дата введена некорректно - выводим окно с ошибкой
    else:
        Analysis(event) # Выполняем функцию Analysis


def Analysis(event): # Функция поиска уязвимостей
    workbook = xlrd2.open_workbook('D:/vullist(xls).xls')
    sheet = workbook.sheet_by_index(0)
    cell = workbook.sheet_by_index(0)

    row = sheet.nrows  # определяем количество записей (строк) на листе
    print('Всего записей', row)  # выведем количество записей на печать

    # выполним считывание списка данных из столбца с данными Название ПО
    names = sheet.col_values(4)  # (4-й столбец, нумерация с нуля)
    status = sheet.col_values(14)
    # выполним считывание списка данных из столбца с данными Уровень опасности
    danger_lavels = sheet.col_values(12)  # (12-й столбец, нумерация с нуля)
    chrb = radioButtonDateVar.get()
    ddd = sheet.col_values(9)

    global danger_low, danger_middle, danger_hight, danger_super
    danger_super, danger_hight, danger_middle, danger_low = 0, 0, 0, 0  # инициализируем переменные-счетчики различных уровней опасности
    if chrb == 0:  # Если радиокнопка По дате выключена (0)
        dataFrom = datetime.strptime('01.01.1900', '%d.%m.%Y')
        dataTo = datetime.strptime('17.06.3021', '%d.%m.%Y')
    else:
        dataFrom = datetime.strptime(textBoxFromDate.get(), '%d.%m.%Y')
        dataTo = datetime.strptime(textBoxToDate.get(), '%d.%m.%Y')

    for i in range(9, row):
        if ddd[i] != '':
            ddd[i] = datetime.strptime(ddd[i], '%d.%m.%Y')
        else:
            ddd[i] = datetime.strptime('01.01.1900', '%d.%m.%Y')

    for i in range(4, row):
        if (str(ddd[i]) >= str(dataFrom)) and (str(ddd[i]) <= str(dataTo)):
            if names[i].find('Adobe Photoshop') >= 0:  # если наименование ПО содержит искомое проверим по первой букве уровень уязвимости ПО
                if danger_lavels[i][0] == 'К':  # Критический
                    danger_super += 1
                elif danger_lavels[i][0] == 'В':  # Высокий
                    danger_hight += 1
                elif danger_lavels[i][0] == 'С':  # Средний
                    danger_middle += 1
                else: # Низкий
                    danger_low += 1

    labelLowOut['text'] = danger_low
    labelMidOut['text'] = danger_middle
    labelHighOut['text'] = danger_hight
    labelSuperOut['text'] = danger_super


def Clear(event): # Функция для очистки лейблов и полей
    labelLowOut['text'] = ""
    labelMidOut['text'] = ""
    labelHighOut['text'] = ""
    labelSuperOut['text'] = ""
    textBoxFromDate.delete(0, END)
    textBoxToDate.delete(0, END)


def SaveDocx(event): # Функция для сохранения результатов в docx
    document = docx.Document()
    document.add_heading('Adobe Photoshop', 0)
    document.add_heading('Количество уязвимостей по уровням опасности', level=1)
    table = document.add_table(rows=4, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '1'
    hdr_cells[1].text = 'Низкий'
    hdr_cells[2].text = str(labelLowOut['text'])
    hdr_cells1 = table.rows[1].cells
    hdr_cells1[0].text = '2'
    hdr_cells1[1].text = 'Средний'
    hdr_cells1[2].text = str(labelMidOut['text'])
    hdr_cells2 = table.rows[2].cells
    hdr_cells2[0].text = '3'
    hdr_cells2[1].text = 'Высокий'
    hdr_cells2[2].text = str(labelHighOut['text'])
    hdr_cells3 = table.rows[3].cells
    hdr_cells3[0].text = '4'
    hdr_cells3[1].text = 'Критический'
    hdr_cells3[2].text = str(labelSuperOut['text'])
    document.save('Анализ уязвимостей Adobe Photoshop.docx')


buttonAnalysis.bind('<Button-1>', AnalysisWithDate) #Привязка функции "AnalysisWithDate" к кнопке "Анализ"
buttonClear.bind('<Button-1>', Clear) #Привязка функции "Clear" к кнопке "Очистить все"
radioButtonDateOff.bind('<Button-1>', dateOff)
radioButtonDateOn.bind('<Button-1>', dateOn) #Привязка функции "dateOn" к радиокнопке "По дате"
buttonSave.bind('<Button-1>', SaveDocx) #Привязка функции "SaveDocx" к кнопке "Сохранить в docx"
labelDate.place(x=120, y=40)
labelDateInfo.pack()
labelFromDate.place(x=130, y=80)
textBoxFromDate.place(x=180, y=80)
labelToDate.place(x=275, y=80)
textBoxToDate.place(x=320, y=80)
labelLow.place(x=30, y=130)
labelLowOut.place(x=180, y=140)
labelMid.place(x=30, y=180)
labelMidOut.place(x=180, y=190)
labelHigh.place(x=30, y=230)
labelHighOut.place(x=180, y=240)
labelSuper.place(x=30, y=280)
labelSuperOut.place(x=180, y=290)
buttonAnalysis.place(x=385, y=130) #Размещаем кнопки по координатам на плоскости окна
buttonSave.place(x=385, y=200) #Размещаем кнопки по координатам на плоскости окна
buttonClear.place(x=385, y=270) #Размещаем кнопки по координатам на плоскости окна
radioButtonDateOn.pack(anchor=W)
radioButtonDateOff.pack(anchor=W)
root.mainloop()

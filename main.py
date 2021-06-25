from tkinter import *
import requests
import docx
import xlrd2
from tkinter import messagebox
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import time
import os.path

root = Tk() # Главное окно
root.title("Welcome") # Название окна
root.geometry("1200x700") # Размер окна
root.configure(background='#f5f5f5') # Цвет заднего фона окна
# Создание интерфейса
radioButtonDateVar = BooleanVar() # Создание радиокнопок
radioButtonDateVar.set(0)
radioButtonDateOn = Radiobutton(text="По дате", bg='#FFFAFA', variable=radioButtonDateVar, value=1)
radioButtonDateOff = Radiobutton(text="За все время", bg='#FFFAFA', variable=radioButtonDateVar, value=0)
# Создание кнопок, полей, лейблов
buttonAnalysis = Button(root, bg='#008B8B', font='Times 12', text="Анализ", width=13, height=2)
buttonClear = Button(root, bg='#008B8B', font='Times 12', text="Удалить", width=13, height=2)
buttonSave = Button(root, bg='#008B8B', font='Times 12', text="Сохранить", width=13, height=2)
buttonSaveEach = Button(root, bg='#008B8B', font='Times 12', text="Сохранить каждый", width=14, height=2)
labelLow = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Низкий")
labelLowOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelMid = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Средний")
labelMidOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelHigh = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Высокий")
labelHighOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelSuper = Label(root, width=13, height=2, bg='#008080', font='Times 13', text="Критический")
labelSuperOut = Label(root, bg='#ffffff', font='Times 15', fg='black', width=5)
labelDate = Label(root, text="Введите необходимую дату:", state=DISABLED,
                  bg='#FFFAFA', font='Times 13', fg='#000', width=30)
labelFromDate = Label(root, text=" От:", state=DISABLED, bg='#FFFAFA', fg='black', width=5)
labelToDate = Label(root, text="До:", state=DISABLED, bg='#FFFAFA', fg='black', width=5)
labelDateInfo = Label(root, text="Анализ уязвимостей Adobe Photoshop", bg='#008080',
                      font='Times 20', fg='#999', width=80)
labelToInfo = Label(root, bg='#FFFAFA', fg='black', width=20)
textBoxFromDate = Entry(root, state=DISABLED, width=10)
textBoxToDate = Entry(root, state=DISABLED, width=10)
buttonDiagram = Button(root, bg='#7fc7ff', font='Times 12', text="Вывести диаграмму", height=2)
buttonobnow = Button(root, bg='#ffd35f', font='Times 12', text="Обновить базу", width=13, height=2)
inputLabel = Label(root, background="violet", font='Times 11', text="Название ПО", width=13)
inputEntry = Entry(root, text='Adobe Photoshop', width=20)
inputEntry.insert(0, "Adobe Photoshop")
operation_label = Label(root, background="#ffbb00", state=DISABLED, text="Статус операции:", width=15)
operation_status_label = Label(root, background="#ffbb00", state=DISABLED, text="Выполняется", width=15)
save_status_label = Label(root, background="#e8594f", state=DISABLED, text="Статус операции", width=15)
operation_save_status_label = Label(root, background="#ffbb00", state=DISABLED, text="Выполняется", width=15)


def Status_executed(event):
    operation_status_label['text'] = "Выполнено"


def download(event):
    operation_label['state'] = NORMAL
    operation_status_label['background'] = "#ffbb00"
    operation_status_label['state'] = NORMAL
    operation_status_label['text'] = "Выполняется"
    files = open('vullist.xlsx', "wb")

    url = 'https://bdu.fstec.ru/files/documents/vullist.xlsx'

    headers = {
        'User-Agent': 'My User Agent 1.0',
        'From': 'youremail@domain.com'  # This is another valid field
    }

    response = requests.get(url, headers=headers)
    files.write(response.content)
    files.close()
    #Status_executed(event)
    operation_status_label['text'] = "Выполнено"


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
    save_status_label['background'] = "#ffbb00"
    save_status_label['state'] = NORMAL
    operation_save_status_label['background'] = "#ffbb00"
    operation_save_status_label['state'] = NORMAL
    operation_save_status_label['text'] = "Выполняется"
    workbook = xlrd2.open_workbook('vullist.xlsx')
    sheet = workbook.sheet_by_index(0)
    cell = workbook.sheet_by_index(0)

    row = sheet.nrows  # определяем количество записей (строк) на листе
    print('Всего записей:', row)  # выведем количество записей на печать

    # выполним считывание списка данных из столбца с данными Название ПО
    names = sheet.col_values(4)  # (4-й столбец, нумерация с нуля)
    status = sheet.col_values(14)
    # выполним считывание списка данных из столбца с данными Уровень опасности
    danger_lavels = sheet.col_values(12)  # (12-й столбец, нумерация с нуля)
    chrb = radioButtonDateVar.get()
    ddd = sheet.col_values(9)
    name_software = inputEntry.get()

    global danger_low, danger_middle, danger_hight, danger_super
    danger_super, danger_hight, danger_middle, danger_low = 0, 0, 0, 0  # инициализируем переменные-счетчики различных
                                                                        # уровней опасности
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
            if names[i].find(name_software) >= 0:  # если наименование ПО содержит искомое проверим по первой
                                                       # букве уровень уязвимости ПО
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
    operation_save_status_label['text'] = "Выполнено"


def Clear(event): # Функция для очистки лейблов и полей
    labelLowOut['text'] = ""
    labelMidOut['text'] = ""
    labelHighOut['text'] = ""
    labelSuperOut['text'] = ""
    operation_status_label['text'] = ""
    textBoxFromDate.delete(0, END)
    textBoxToDate.delete(0, END)


def SaveDocx(event): # Функция для сохранения результатов в docx
    if (labelLowOut['text'] and labelMidOut['text'] and labelHighOut['text'] and labelSuperOut['text']) == "":
        messagebox.showerror("Error", "Сначала нужно провести анализ данных!")
    else:
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


def SaveDocxEach(event):
    list_var = ["низкому", "среднему", "высокому", "критическому"]
    list_var_1 = ["низких", "средних", "высоких", "критических"]
    list_var_2 = ["Низкий", "Средний", "Высокий", "Критический"]

    for i in range(4):
        document = docx.Document()
        document.add_heading('Adobe Photoshop', 0)
        document.add_heading('Количество уязвимостей по {} уровню опасности '.format(list_var[i]), level=1)
        table = document.add_table(rows=1, cols=3)

        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '1'
        hdr_cells[1].text = list_var_2[i]

        if i == 0:
            hdr_cells[2].text = str(labelLowOut['text'])
        elif i == 1:
            hdr_cells[2].text = str(labelMidOut['text'])
        elif i == 2:
            hdr_cells[2].text = str(labelHighOut['text'])
        else:
            hdr_cells[2].text = str(labelSuperOut['text'])

        document.save('Анализ {} уязвимостей Adobe Photoshop.docx'.format(list_var_1[i]))


def diagramma(event):
    try:
        print(danger_low)
        labels = 'Низкий', 'Средний', 'Высокий', 'Критический'
        sizes = [danger_low, danger_middle, danger_hight, danger_super]

        colors = ("grey", "yellow", "orange", "brown")
        fig1, ax1 = plt.subplots()
        explode = (0, 0, 0.1, 0.15)

        ax1.pie(sizes, colors=colors, explode=explode, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)
        patches, texts, auto = ax1.pie(sizes, colors=colors, shadow=True, startangle=90, explode=explode, autopct='%1.1f%%' )

        plt.legend(patches, labels, loc="best")
        window = Tk()
        window.title("Диаграмма уязвимостей")
        window.configure(background='#a8e4a0')
        canvas = FigureCanvasTkAgg(fig1, master=window)
        canvas.get_tk_widget().pack()
        canvas.draw()
    except NameError:
        messagebox.showerror("Error", "Для вывода диаграммы необходимо провести анализ!")


buttonAnalysis.bind('<Button-1>', AnalysisWithDate) #Привязка функции "AnalysisWithDate" к кнопке "Анализ"
buttonClear.bind('<Button-1>', Clear) #Привязка функции "Clear" к кнопке "Очистить все"
radioButtonDateOff.bind('<Button-1>', dateOff)
radioButtonDateOn.bind('<Button-1>', dateOn) #Привязка функции "dateOn" к радиокнопке "По дате"
buttonSave.bind('<Button-1>', SaveDocx) #Привязка функции "SaveDocx" к кнопке "Сохранить в docx"
buttonSaveEach.bind('<Button-1>', SaveDocxEach)
buttonDiagram.bind('<Button-1>', diagramma)
buttonobnow.bind('<Button-1>', download)
#inputEntry.bind('<Button-1>', dd)

buttonobnow.place(x=441, y=60)
buttonDiagram.place(x=250, y=130)
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
buttonAnalysis.place(x=440, y=130) #Размещаем кнопки по координатам на плоскости окна
buttonSave.place(x=440, y=200) #Размещаем кнопки по координатам на плоскости окна
buttonSaveEach.place(x=580, y=200) #Размещаем кнопки по координатам на плоскости окна
buttonClear.place(x=440, y=270) #Размещаем кнопки по координатам на плоскости окна
inputLabel.place(x=100, y=105)
inputEntry.place(x=270, y=105)
radioButtonDateOn.pack(anchor=W)
radioButtonDateOff.pack(anchor=W)
operation_label.place(x=600, y=63)
operation_status_label.place(x=600, y=87)
save_status_label.place(x=600, y=133)
operation_save_status_label.place(x=600, y=157)
#print(os.path.exists('vullist.xlsx'))
print("Круг программы выполнен")
root.mainloop()

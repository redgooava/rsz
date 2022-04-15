import tkinter
from tkinter.ttk import Combobox

import psycopg2
from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar, DateEntry
import datetime

con = psycopg2.connect(
    database="postgres",
    user="postgres",
    password="",
    host="127.0.0.1",
    port="5432"
)
cur = con.cursor()

# окошко
window = Tk()
window.geometry('1000x750')
window.title("Строевая записка")

# менюбар
tab_control = ttk.Notebook(window)
tab1 = ttk.Frame(tab_control)
tab2 = ttk.Frame(tab_control)
tab3 = ttk.Frame(tab_control)
tab4 = ttk.Frame(tab_control)
tab5 = ttk.Frame(tab_control)
tab6 = ttk.Frame(tab_control)
tab7 = ttk.Frame(tab_control)
tab8 = ttk.Frame(tab_control)
tab9 = ttk.Frame(tab_control)
tab10 = ttk.Frame(tab_control)
tab11 = ttk.Frame(tab_control)
tab_control.add(tab1, text='Таблица')
tab_control.add(tab2, text='Отпуск')
tab_control.add(tab3, text='Командировка')
tab_control.add(tab4, text='На амбулаторном лечении')
tab_control.add(tab5, text='Наряд')
tab_control.add(tab6, text='Госпиталь')
tab_control.add(tab7, text='Поиск человека')
tab_control.add(tab8, text='Поиск подразделения')
tab_control.add(tab9, text='По штату/списку')
tab_control.add(tab10, text='Правка штата')
tab_control.add(tab11, text='О программе')
tab_control.pack(expand=1, fill='both')

actionRows = []

listOfDivisions = ('Стелс-пехота', 'Хлеборезы', 'Дикий стройбат', 'Штабные убийцы')
listOfRanks = ('Оф', 'Пр', 'К/с', 'С/с', 'К-ты', 'Сл')


# добавление записи (для отпуска, командировки и прочих записей об отсутствии)
def actionInput(name, whyout, start, end, sDivision, numberOfOrder, rank):
    if (name is not None) and (whyout is not None) and (start is not None) and (end is not None):
        cur.execute(
            f"INSERT INTO OUTTABLE (NAME, WHYOUT, DATESTART, DATEEND, SUBDIVISION, NUMBEROFORDER, RANK) "
            f"VALUES ('{name}', '{whyout}', '{start}', '{end}', '{sDivision}', '{numberOfOrder}', '{rank}')"
        )
        con.commit()
    else:
        print('Ошибка в actionInput')


def actionSearch(name, table):
    cur.execute(
        f"SELECT * FROM OUTTABLE "
        f"WHERE name = '{name}'"
    )
    for item in table.get_children():
        table.delete(item)

    actionRows = cur.fetchall()
    con.commit()

    for row in actionRows:
        table.insert('', tkinter.END, values=row)

    table.pack(expand=tkinter.YES, fill=tkinter.BOTH)
    table.place(x=0, y=200)


def actionSearchDivision(division, table):
    cur.execute(
        f"SELECT * FROM OUTTABLE "
        f"WHERE subdivision = '{division}'"
    )

    for item in table.get_children():
        table.delete(item)

    actionRows = cur.fetchall()
    con.commit()

    for row in actionRows:
        table.insert('', tkinter.END, values=row)

    table.pack(expand=tkinter.YES, fill=tkinter.BOTH)
    table.place(x=0, y=200)


def actionEdit(division, officer, ensign, contract, soldier, cadet, listener, total,
               s_officer, s_ensign, s_contract, s_soldier, s_cadet, s_listener, s_total):
    cur.execute(
        f"INSERT INTO DIVISIONTABLE (DIVISION, OFFICER, ENSIGN, CONTRACT, SOLDIER, CADET, LISTENER, TOTAL, "
        f" SOFFICER, SENSIGN, SCONTRACT, SSOLDIER, SCADET, SLISTENER, STOTAL)"
        f"VALUES('{division}', '{officer}', '{ensign}', '{contract}', '{soldier}', '{cadet}', '{listener}', '{total}', "
        f"'{s_officer}', '{s_ensign}', '{s_contract}', '{s_soldier}', '{s_cadet}', '{s_listener}', '{s_total}')"
    )
    con.commit()


def actionGeneral(table):
    cur.execute(
        f"SELECT * FROM DIVISIONTABLE"
    )
    con.commit()

    for item in table.get_children():
        table.delete(item)

    actionRows = cur.fetchall()

    for i in range(len(actionRows)):
        actionRows[i] = list(actionRows[i])
        actionRows[i].append(actionRows[i][9] - inPeriodAtDivision(actionRows[i][1], 'Оф'))
        actionRows[i].append(actionRows[i][10] - inPeriodAtDivision(actionRows[i][1], 'Пр'))
        actionRows[i].append(actionRows[i][11] - inPeriodAtDivision(actionRows[i][1], 'К/c'))
        actionRows[i].append(actionRows[i][12] - inPeriodAtDivision(actionRows[i][1], 'C/c'))
        actionRows[i].append(actionRows[i][13] - inPeriodAtDivision(actionRows[i][1], 'К-ты'))
        actionRows[i].append(actionRows[i][14] - inPeriodAtDivision(actionRows[i][1], 'Сл'))
        actionRows[i].append(actionRows[i][16] + actionRows[i][17] + actionRows[i][18] +
                             actionRows[i][19] + actionRows[i][20] + actionRows[i][21])
        actionRows[i].append(inPeriodNotAtDivision(actionRows[i][1], 'Отпуск'))
        actionRows[i].append(inPeriodNotAtDivision(actionRows[i][1], 'Госпиталь'))
        actionRows[i].append(inPeriodNotAtDivision(actionRows[i][1], 'На амбулаторном лечении'))
        actionRows[i].append(inPeriodNotAtDivision(actionRows[i][1], 'В наряде'))
        actionRows[i].append(inPeriodNotAtDivision(actionRows[i][1], 'Командировка'))
        actionRows[i].append(actionRows[i][23] + actionRows[i][24] +
                             actionRows[i][25] + actionRows[i][26] + actionRows[i][27])

    con.commit()

    for row in actionRows:
        table.insert('', tkinter.END, values=row)

    table.pack(expand=tkinter.YES, fill=tkinter.BOTH)
    table.place(x=40, y=50, width=900, height=400)


def inPeriodAtDivision(division, rank):
    cur.execute(
        f"SELECT * FROM OUTTABLE "
        f"WHERE subdivision = '{division}' AND (rank = '{rank}') AND "
        f"(datestart <= now()) AND "
        f"(dateend > now())"
    )
    con.commit()

    actionRows = cur.fetchall()
    return len(actionRows)


def inPeriodNotAtDivision(division, whyout):
    cur.execute(
        f"SELECT * FROM OUTTABLE "
        f"WHERE subdivision = '{division}' AND whyout = '{whyout}' AND "
        f"(datestart <= now()) AND "
        f"(dateend > now())"
    )
    con.commit()

    actionRows = cur.fetchall()
    return len(actionRows)


# вкладка "Командировка"
def trip():
    general_lbl = Label(tab3, text='Командировка', font=("Arial", 20))
    general_lbl.place(width=250, height=30, x=130, y=70)
    name_trip_lbl = Label(tab3, text="ФИО: ")
    name_trip_lbl.place(width=250, height=30, x=0, y=150)
    name_trip_ent = Entry(tab3, width=10, font=("Verdana", 12))
    name_trip_ent.place(width=250, height=30, x=180, y=150)
    cal_start_trip_lbl = Label(tab3, text="Дата убытия: ")
    cal_start_trip_lbl.place(width=250, height=30, x=0, y=200)
    cal_start_trip = DateEntry(tab3, width=16, background="magenta3", foreground="white", bd=2)
    cal_start_trip.pack(pady=20)
    cal_start_trip.place(width=250, height=30, x=180, y=200)
    cal_end_trip_lbl = Label(tab3, text="Дата прибытия: ")
    cal_end_trip_lbl.place(width=250, height=30, x=0, y=250)
    cal_end_trip = DateEntry(tab3, width=16, background="magenta3", foreground="white", bd=2)
    cal_end_trip.pack(pady=20)
    cal_end_trip.place(width=250, height=30, x=180, y=250)
    division_lbl = Label(tab3, text="Подразделение: ")
    division_lbl.place(width=250, height=30, x=0, y=300)
    division = Combobox(tab3, values=listOfDivisions)
    division.place(width=250, height=30, x=180, y=300)
    numberOfOrder_lbl = Label(tab3, text="Номер приказа: ")
    numberOfOrder_lbl.place(width=250, height=30, x=0, y=350)
    numberOfOrder = Entry(tab3, width=10, font=("Verdana", 12))
    numberOfOrder.place(width=250, height=30, x=180, y=350)
    rank_lbl = Label(tab3, text="Категория: ")
    rank_lbl.place(width=250, height=30, x=0, y=400)
    rank = Combobox(tab3, values=listOfRanks)
    rank.place(width=250, height=30, x=180, y=400)
    button_trip = Button(tab3, text='Записать', command=lambda: actionInput(
        name_trip_ent.get(), 'Командировка',
        cal_start_trip.get_date().strftime("%d/%m/%Y"), cal_end_trip.get_date().strftime("%d/%m/%Y"), division.get(),
        numberOfOrder.get(), rank.get()))
    button_trip.place(width=100, height=30, x=330, y=450)


# вкладка "Отпуск"
def vacation():
    general_lbl = Label(tab2, text='Отпуск', font=("Arial", 20))
    general_lbl.place(width=250, height=30, x=130, y=70)
    name_trip_lbl = Label(tab2, text="ФИО: ")
    name_trip_lbl.place(width=250, height=30, x=0, y=150)
    name_trip_ent = Entry(tab2, width=10, font=("Verdana", 12))
    name_trip_ent.place(width=250, height=30, x=180, y=150)
    cal_start_trip_lbl = Label(tab2, text="Дата убытия: ")
    cal_start_trip_lbl.place(width=250, height=30, x=0, y=200)
    cal_start_trip = DateEntry(tab2, width=16, background="magenta3", foreground="white", bd=2)
    cal_start_trip.pack(pady=20)
    cal_start_trip.place(width=250, height=30, x=180, y=200)
    cal_end_trip_lbl = Label(tab2, text="Дата прибытия: ")
    cal_end_trip_lbl.place(width=250, height=30, x=0, y=250)
    cal_end_trip = DateEntry(tab2, width=16, background="magenta3", foreground="white", bd=2)
    cal_end_trip.pack(pady=20)
    cal_end_trip.place(width=250, height=30, x=180, y=250)
    division_lbl = Label(tab2, text="Подразделение: ")
    division_lbl.place(width=250, height=30, x=0, y=300)
    division = Combobox(tab2, values=listOfDivisions)
    division.place(width=250, height=30, x=180, y=300)
    numberOfOrder_lbl = Label(tab2, text="Номер приказа: ")
    numberOfOrder_lbl.place(width=250, height=30, x=0, y=350)
    numberOfOrder = Entry(tab2, width=10, font=("Verdana", 12))
    numberOfOrder.place(width=250, height=30, x=180, y=350)
    rank_lbl = Label(tab2, text="Категория: ")
    rank_lbl.place(width=250, height=30, x=0, y=400)
    rank = Combobox(tab2, values=listOfRanks)
    rank.place(width=250, height=30, x=180, y=400)
    button_trip = Button(tab2, text='Записать', command=lambda: actionInput(
        name_trip_ent.get(), 'Отпуск',
        cal_start_trip.get_date().strftime("%d/%m/%Y"), cal_end_trip.get_date().strftime("%d/%m/%Y"), division.get(),
        numberOfOrder.get(), rank.get()))
    button_trip.place(width=100, height=30, x=330, y=450)


# вкладка "На амбулаторном лечении"
def therapy():
    general_lbl = Label(tab4, text='На амбулаторном лечении', font=("Arial", 20))
    general_lbl.place(width=350, height=30, x=100, y=70)
    name_trip_lbl = Label(tab4, text="ФИО: ")
    name_trip_lbl.place(width=250, height=30, x=0, y=150)
    name_trip_ent = Entry(tab4, width=10, font=("Verdana", 12))
    name_trip_ent.place(width=250, height=30, x=180, y=150)
    cal_start_trip_lbl = Label(tab4, text="Дата убытия: ")
    cal_start_trip_lbl.place(width=250, height=30, x=0, y=200)
    cal_start_trip = DateEntry(tab4, width=16, background="magenta3", foreground="white", bd=2)
    cal_start_trip.pack(pady=20)
    cal_start_trip.place(width=250, height=30, x=180, y=200)
    cal_end_trip_lbl = Label(tab4, text="Дата прибытия: ")
    cal_end_trip_lbl.place(width=250, height=30, x=0, y=250)
    cal_end_trip = DateEntry(tab4, width=16, background="magenta3", foreground="white", bd=2)
    cal_end_trip.pack(pady=20)
    cal_end_trip.place(width=250, height=30, x=180, y=250)
    division_lbl = Label(tab4, text="Подразделение: ")
    division_lbl.place(width=250, height=30, x=0, y=300)
    division = Combobox(tab4, values=listOfDivisions)
    division.place(width=250, height=30, x=180, y=300)
    numberOfOrder_lbl = Label(tab4, text="Номер приказа: ")
    numberOfOrder_lbl.place(width=250, height=30, x=0, y=350)
    numberOfOrder = Entry(tab4, width=10, font=("Verdana", 12))
    numberOfOrder.place(width=250, height=30, x=180, y=350)
    rank_lbl = Label(tab4, text="Категория: ")
    rank_lbl.place(width=250, height=30, x=0, y=400)
    rank = Combobox(tab4, values=listOfRanks)
    rank.place(width=250, height=30, x=180, y=400)
    button_trip = Button(tab4, text='Записать', command=lambda: actionInput(
        name_trip_ent.get(), 'На амбулаторном лечении',
        cal_start_trip.get_date().strftime("%d/%m/%Y"), cal_end_trip.get_date().strftime("%d/%m/%Y"), division.get(),
        numberOfOrder.get(), rank.get()))
    button_trip.place(width=100, height=30, x=330, y=450)


# вкладка "В наряде"
def tourOfDuty():
    general_lbl = Label(tab5, text='В наряде', font=("Arial", 20))
    general_lbl.place(width=250, height=30, x=130, y=70)
    name_trip_lbl = Label(tab5, text="ФИО: ")
    name_trip_lbl.place(width=250, height=30, x=0, y=150)
    name_trip_ent = Entry(tab5, width=10, font=("Verdana", 12))
    name_trip_ent.place(width=250, height=30, x=180, y=150)
    cal_start_trip_lbl = Label(tab5, text="Дата убытия: ")
    cal_start_trip_lbl.place(width=250, height=30, x=0, y=200)
    cal_start_trip = DateEntry(tab5, width=16, background="magenta3", foreground="white", bd=2)
    cal_start_trip.pack(pady=20)
    cal_start_trip.place(width=250, height=30, x=180, y=200)
    cal_end_trip_lbl = Label(tab5, text="Дата прибытия: ")
    cal_end_trip_lbl.place(width=250, height=30, x=0, y=250)
    cal_end_trip = DateEntry(tab5, width=16, background="magenta3", foreground="white", bd=2)
    cal_end_trip.pack(pady=20)
    cal_end_trip.place(width=250, height=30, x=180, y=250)
    division_lbl = Label(tab5, text="Подразделение: ")
    division_lbl.place(width=250, height=30, x=0, y=300)
    division = Combobox(tab5, values=listOfDivisions)
    division.place(width=250, height=30, x=180, y=300)
    numberOfOrder_lbl = Label(tab5, text="Номер приказа: ")
    numberOfOrder_lbl.place(width=250, height=30, x=0, y=350)
    numberOfOrder = Entry(tab5, width=10, font=("Verdana", 12))
    numberOfOrder.place(width=250, height=30, x=180, y=350)
    rank_lbl = Label(tab5, text="Категория: ")
    rank_lbl.place(width=250, height=30, x=0, y=400)
    rank = Combobox(tab5, values=listOfRanks)
    rank.place(width=250, height=30, x=180, y=400)
    button_trip = Button(tab5, text='Записать', command=lambda: actionInput(
        name_trip_ent.get(), 'В наряде',
        cal_start_trip.get_date().strftime("%d/%m/%Y"), cal_end_trip.get_date().strftime("%d/%m/%Y"), division.get(),
        numberOfOrder.get(), rank.get()))
    button_trip.place(width=100, height=30, x=330, y=450)


# вкладка "В наряде"
def hospital():
    general_lbl = Label(tab6, text='Госпиталь', font=("Arial", 20))
    general_lbl.place(width=250, height=30, x=130, y=70)
    name_trip_lbl = Label(tab6, text="ФИО: ")
    name_trip_lbl.place(width=250, height=30, x=0, y=150)
    name_trip_ent = Entry(tab6, width=10, font=("Verdana", 12))
    name_trip_ent.place(width=250, height=30, x=180, y=150)
    cal_start_trip_lbl = Label(tab6, text="Дата убытия: ")
    cal_start_trip_lbl.place(width=250, height=30, x=0, y=200)
    cal_start_trip = DateEntry(tab6, width=16, background="magenta3", foreground="white", bd=2)
    cal_start_trip.pack(pady=20)
    cal_start_trip.place(width=250, height=30, x=180, y=200)
    cal_end_trip_lbl = Label(tab6, text="Дата прибытия: ")
    cal_end_trip_lbl.place(width=250, height=30, x=0, y=250)
    cal_end_trip = DateEntry(tab6, width=16, background="magenta3", foreground="white", bd=2)
    cal_end_trip.pack(pady=20)
    cal_end_trip.place(width=250, height=30, x=180, y=250)
    division_lbl = Label(tab6, text="Подразделение: ")
    division_lbl.place(width=250, height=30, x=0, y=300)
    division = Combobox(tab6, values=listOfDivisions)
    division.place(width=250, height=30, x=180, y=300)
    numberOfOrder_lbl = Label(tab6, text="Номер приказа: ")
    numberOfOrder_lbl.place(width=250, height=30, x=0, y=350)
    numberOfOrder = Entry(tab6, width=10, font=("Verdana", 12))
    numberOfOrder.place(width=250, height=30, x=180, y=350)
    rank_lbl = Label(tab6, text="Категория: ")
    rank_lbl.place(width=250, height=30, x=0, y=400)
    rank = Combobox(tab6, values=listOfRanks)
    rank.place(width=250, height=30, x=180, y=400)
    button_trip = Button(tab6, text='Записать', command=lambda: actionInput(
        name_trip_ent.get(), 'Госпиталь',
        cal_start_trip.get_date().strftime("%d/%m/%Y"), cal_end_trip.get_date().strftime("%d/%m/%Y"), division.get(),
        numberOfOrder.get(), rank.get()))
    button_trip.place(width=100, height=30, x=330, y=450)


# вкладка "Поиск человека"
def search():
    name_search_lbl = Label(tab7, text="ФИО: ")
    name_search_lbl.place(width=250, height=30, x=0, y=50)
    name_search_ent = Entry(tab7, width=10, font=("Verdana", 12))
    name_search_ent.place(width=250, height=30, x=180, y=50)

    # всё про таблицу отсюда https://www.youtube.com/watch?v=HMPIeZ3S_cs
    table = ttk.Treeview(tab7, show='headings')
    table.place(x=0, y=200)
    heads = ['ФИО', 'Причина отсутствия', 'C какого дня', 'По какой день', 'Подразделение', 'Номера приказа']
    table['columns'] = heads

    for header in heads:
        table.heading(header, text=header, anchor='center')
        table.column(header, anchor='center', width=200)

    scroll_pane = ttk.Scrollbar(tab7, command=table.yview)
    table.configure(yscrollcommand=scroll_pane.set)
    scroll_pane.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    scroll_pane.place(height=220, x=1204, y=200)

    button_search = Button(tab7, text='Найти', command=lambda: actionSearch(name_search_ent.get(), table))
    button_search.place(width=100, height=30, x=450, y=50)


def searchDivision():
    division_search_lbl = Label(tab8, text="Подразделение: ")
    division_search_lbl.place(width=250, height=30, x=0, y=50)
    division_search_ent = Entry(tab8, width=10, font=("Verdana", 12))
    division_search_ent.place(width=250, height=30, x=180, y=50)

    # всё про таблицу отсюда https://www.youtube.com/watch?v=HMPIeZ3S_cs
    dtable = ttk.Treeview(tab8, show='headings')
    dtable.place(x=0, y=200)
    heads = ['ФИО', 'Причина отсутствия', 'C какого дня', 'По какой день', 'Подразделение', 'Номера приказа']
    dtable['columns'] = heads

    for header in heads:
        dtable.heading(header, text=header, anchor='center')
        dtable.column(header, anchor='center', width=200)

    scroll_pane = ttk.Scrollbar(tab8, command=dtable.yview)
    dtable.configure(yscrollcommand=scroll_pane.set)
    scroll_pane.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    scroll_pane.place(height=220, x=1204, y=200)

    button_search = Button(tab8, text='Найти', command=lambda: actionSearchDivision(division_search_ent.get(), dtable))
    button_search.place(width=100, height=30, x=450, y=50)


def defaultList():
    columns = ("#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10", "#11", "#12", "#13", "#14", "#15", "#16")
    table = ttk.Treeview(tab9, show="headings", columns=columns)
    table.place(x=0, y=200)
    table.heading("#1", text="№п/п")
    table.column("#1", anchor='center', width=50)
    table.heading("#2", text="Подразделение")
    table.column("#2", anchor='center', width=100)
    table.heading("#3", text="Оф")
    table.column("#3", anchor='center', width=50)
    table.heading("#4", text="Пр")
    table.column("#4", anchor='center', width=50)
    table.heading("#5", text="К/с")
    table.column("#5", anchor='center', width=50)
    table.heading("#6", text="С/с")
    table.column("#6", anchor='center', width=50)
    table.heading("#7", text="К-ты")
    table.column("#7", anchor='center', width=50)
    table.heading("#8", text="Сл")
    table.column("#8", anchor='center', width=50)
    table.heading("#9", text="ВСЕГО")
    table.column("#9", anchor='center', width=50)
    table.heading("#10", text="Оф")
    table.column("#10", anchor='center', width=50)
    table.heading("#11", text="Пр")
    table.column("#11", anchor='center', width=50)
    table.heading("#12", text="К/с")
    table.column("#12", anchor='center', width=50)
    table.heading("#13", text="С/с")
    table.column("#13", anchor='center', width=50)
    table.heading("#14", text="К-ты")
    table.column("#14", anchor='center', width=50)
    table.heading("#15", text="Сл")
    table.column("#15", anchor='center', width=50)
    table.heading("#16", text="ВСЕГО")
    table.column("#16", anchor='center', width=50)

    scroll_pane = ttk.Scrollbar(tab9, command=table.yview)
    table.configure(yscrollcommand=scroll_pane.set)
    scroll_pane.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    scroll_pane.place(height=220, x=1204, y=200)


def generalTable():
    columns = ("#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10", "#11", "#12", "#13", "#14", "#15", "#16",
               "#17", "#18", "#19", "#20", "#21", "#22", "#23", "#24", "#25", "#26", "#27", "#28", "#29")
    table = ttk.Treeview(tab1, show="headings", columns=columns)
    table.place(x=40, y=50, width=900, height=400)
    table.heading("#1", text="№п/п")
    table.column("#1", anchor='center', width=40)
    table.heading("#2", text="Подразделение")
    table.column("#2", anchor='center', width=120)
    table.heading("#3", text="Оф")
    table.column("#3", anchor='center', width=50)
    table.heading("#4", text="Пр")
    table.column("#4", anchor='center', width=50)
    table.heading("#5", text="К/с")
    table.column("#5", anchor='center', width=50)
    table.heading("#6", text="С/с")
    table.column("#6", anchor='center', width=50)
    table.heading("#7", text="К-ты")
    table.column("#7", anchor='center', width=50)
    table.heading("#8", text="Сл")
    table.column("#8", anchor='center', width=50)
    table.heading("#9", text="ВСЕГО")
    table.column("#9", anchor='center', width=50)
    table.heading("#10", text="Оф")
    table.column("#10", anchor='center', width=50)
    table.heading("#11", text="Пр")
    table.column("#11", anchor='center', width=50)
    table.heading("#12", text="К/с")
    table.column("#12", anchor='center', width=50)
    table.heading("#13", text="С/с")
    table.column("#13", anchor='center', width=50)
    table.heading("#14", text="К-ты")
    table.column("#14", anchor='center', width=50)
    table.heading("#15", text="Сл")
    table.column("#15", anchor='center', width=50)
    table.heading("#16", text="ВСЕГО")
    table.column("#16", anchor='center', width=50)
    table.heading("#17", text="Оф")
    table.column("#17", anchor='center', width=50)
    table.heading("#18", text="Пр")
    table.column("#18", anchor='center', width=50)
    table.heading("#19", text="К/с")
    table.column("#19", anchor='center', width=50)
    table.heading("#20", text="С/с")
    table.column("#20", anchor='center', width=50)
    table.heading("#21", text="К-ты")
    table.column("#21", anchor='center', width=50)
    table.heading("#22", text="Сл")
    table.column("#22", anchor='center', width=50)
    table.heading("#23", text="ВСЕГО")
    table.column("#23", anchor='center', width=50)
    table.heading("#24", text="Отпуск")
    table.column("#24", anchor='center', width=50)
    table.heading("#25", text="Стац")
    table.column("#25", anchor='center', width=50)
    table.heading("#26", text="Амб")
    table.column("#26", anchor='center', width=50)
    table.heading("#27", text="Наряд")
    table.column("#27", anchor='center', width=50)
    table.heading("#28", text="Ком-ка")
    table.column("#28", anchor='center', width=50)
    table.heading("#29", text="ВСЕГО")
    table.column("#29", anchor='center', width=50)

    scroll_pane = ttk.Scrollbar(tab1, command=table.yview)
    table.configure(yscrollcommand=scroll_pane.set)
    scroll_pane.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    scroll_pane.place(height=400, x=940, y=50)

    scroll_pane1 = ttk.Scrollbar(tab1, command=table.xview, orient='horizontal')
    table.configure(xscrollcommand=scroll_pane1.set)
    scroll_pane1.pack(side=tkinter.BOTTOM, fill=tkinter.X)
    scroll_pane1.place(width=900, x=40, y=450)

    button_general = Button(tab1, text='Сформировать', command=lambda: actionGeneral(table))
    button_general.place(width=100, height=30, x=820, y=490)


def editDivisionTable():
    general_lbl = Label(tab10, text='Добавление подразделения', font=("Arial", 20))
    general_lbl.place(width=350, height=30, x=100, y=70)

    division_lbl = Label(tab10, text="Подразделение: ")
    division_lbl.place(width=250, height=30, x=0, y=150)
    division_ent = Entry(tab10, width=10, font=("Verdana", 12))
    division_ent.place(width=250, height=30, x=180, y=150)

    officer_lbl = Label(tab10, text="Количество офицеров: ")
    officer_lbl.place(width=250, height=30, x=0, y=200)
    officer_ent = Entry(tab10, width=10, font=("Verdana", 12))
    officer_ent.place(width=250, height=30, x=180, y=200)

    ensign_lbl = Label(tab10, text="Количество прапорщиков: ")
    ensign_lbl.place(width=250, height=30, x=0, y=250)
    ensign_ent = Entry(tab10, width=10, font=("Verdana", 12))
    ensign_ent.place(width=250, height=30, x=180, y=250)

    contract_lbl = Label(tab10, text="Количество в/с к/с: ")
    contract_lbl.place(width=250, height=30, x=0, y=300)
    contract_ent = Entry(tab10, width=10, font=("Verdana", 12))
    contract_ent.place(width=250, height=30, x=180, y=300)

    soldier_lbl = Label(tab10, text="Количество  в/с с/с: ")
    soldier_lbl.place(width=250, height=30, x=0, y=350)
    soldier_ent = Entry(tab10, width=10, font=("Verdana", 12))
    soldier_ent.place(width=250, height=30, x=180, y=350)

    cadet_lbl = Label(tab10, text="Количество курсантов: ")
    cadet_lbl.place(width=250, height=30, x=0, y=400)
    cadet_ent = Entry(tab10, width=10, font=("Verdana", 12))
    cadet_ent.place(width=250, height=30, x=180, y=400)

    listener_lbl = Label(tab10, text="Количество слушателей: ")
    listener_lbl.place(width=250, height=30, x=0, y=450)
    listener_ent = Entry(tab10, width=10, font=("Verdana", 12))
    listener_ent.place(width=250, height=30, x=180, y=450)

    total_lbl = Label(tab10, text="Общее количество: ")
    total_lbl.place(width=250, height=30, x=0, y=500)
    total_ent = Entry(tab10, width=10, font=("Verdana", 12))
    total_ent.place(width=250, height=30, x=180, y=500)

    s_officer_lbl = Label(tab10, text="Количество офицеров: ")
    s_officer_lbl.place(width=250, height=30, x=400, y=200)
    s_officer_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_officer_ent.place(width=250, height=30, x=580, y=200)

    s_ensign_lbl = Label(tab10, text="Количество прапорщиков: ")
    s_ensign_lbl.place(width=250, height=30, x=400, y=250)
    s_ensign_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_ensign_ent.place(width=250, height=30, x=580, y=250)

    s_contract_lbl = Label(tab10, text="Количество в/с к/с: ")
    s_contract_lbl.place(width=250, height=30, x=400, y=300)
    s_contract_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_contract_ent.place(width=250, height=30, x=580, y=300)

    s_soldier_lbl = Label(tab10, text="Количество  в/с с/с: ")
    s_soldier_lbl.place(width=250, height=30, x=400, y=350)
    s_soldier_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_soldier_ent.place(width=250, height=30, x=580, y=350)

    s_cadet_lbl = Label(tab10, text="Количество курсантов: ")
    s_cadet_lbl.place(width=250, height=30, x=400, y=400)
    s_cadet_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_cadet_ent.place(width=250, height=30, x=580, y=400)

    s_listener_lbl = Label(tab10, text="Количество слушателей: ")
    s_listener_lbl.place(width=250, height=30, x=400, y=450)
    s_listener_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_listener_ent.place(width=250, height=30, x=580, y=450)

    s_total_lbl = Label(tab10, text="Общее количество: ")
    s_total_lbl.place(width=250, height=30, x=400, y=500)
    s_total_ent = Entry(tab10, width=10, font=("Verdana", 12))
    s_total_ent.place(width=250, height=30, x=580, y=500)

    button_division = Button(tab10, text='Добавить', command=lambda: actionEdit(
        division_ent.get(), officer_ent.get(), ensign_ent.get(), contract_ent.get(),
        soldier_ent.get(), cadet_ent.get(), listener_ent.get(), total_ent.get(),
        s_officer_ent.get(), s_ensign_ent.get(), s_contract_ent.get(),
        s_soldier_ent.get(), s_cadet_ent.get(), s_listener_ent.get(), s_total_ent.get()
    ))
    button_division.place(width=100, height=30, x=330, y=550)


trip()
vacation()
search()
therapy()
tourOfDuty()
hospital()
searchDivision()
defaultList()
generalTable()
editDivisionTable()

window.mainloop()

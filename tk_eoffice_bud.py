#!python3
"""
Ten plik pyta w GUI o zakres plików testi do przerobienia - pierwszy numer pliku i ostatni.
Wrzuca do osobnych arkuszy i do arkusza głównego.
Plik poprzedni: tk_eoffice_bud_old.py wrzuca dane do jednego arkusza
"""
#How to Open and Close Programs: https://www.youtube.com/watch?v=IJm9m8kv7gU
'''=========================
nie działa, nie pobiera pierwszy_tekst do funkcji eoffice_et. Zapytać Mateusza.
============================='''

#https://www.tutorialspoint.com/python/python_gui_programming.htm
#https://docs.python.org/3.6/library/tkinter.html
#https://www.python-course.eu/tkinter_entry_widgets.php
from functions_file_new import create_file, open_file, eoffice_et

create_file()

import tkinter
from tkinter import *
czcionka = "Verdana 11 italic"
tło = "LightBlue1"
kolor_czcionki = "LightSteelBlue4"
root = Tk()
root.minsize(600,100)
root.title("Pobieranie danych z IR do pliku budżetowego")
root.configure(background = tło)

topframe = Frame(root)
topframe.pack()
pierwszy = Label(topframe, text="Wprowadź trzycyfrowy numer pierwszego pliku testi:", font = czcionka, bg = tło, fg = kolor_czcionki)
pierwszy_tekst = Entry(topframe, bd = 2)
pierwszy.pack(side = LEFT)
pierwszy_tekst.pack(side = RIGHT)
pierwszy_tekst.focus_set()
#focus_set umieszcza kursor w pierwszym okienku

midframe = Frame(root)
midframe.pack(side = TOP)
ostatni = Label(midframe, text="Wprowadź trzycyfrowy numer ostatniego pliku testi:", font = czcionka, bg = tło, fg = kolor_czcionki)
ostatni_tekst = Entry(midframe, bd = 2)
ostatni.pack(side = LEFT)
ostatni_tekst.pack(side = RIGHT)

bottomframe = Frame(root)
bottomframe.pack(side = BOTTOM)
runButton = Button(bottomframe, text ="Uruchom", font = czcionka, fg = kolor_czcionki, command = eoffice_et, bd = 2).pack(side = LEFT)
exitButton = Button(bottomframe, text = "Zakończ", font = czcionka, fg = kolor_czcionki, command = root.destroy, bd = 2).pack(side = RIGHT)
openButton = Button(bottomframe, text = "Otwórz plik", font = czcionka, fg = kolor_czcionki, command = open_file, bd = 2).pack(side = RIGHT)

#bd = border
mainloop()

#https://www.youtube.com/watch?v=tMhFk_GgmFk

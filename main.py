import eel
from tkinter import filedialog
import tkinter as tk
import pandas
from back.main_back import process
import os

dir_path = os.path.dirname(os.path.realpath(__file__))

eel.init(dir_path)


@eel.expose
def choose_file():
    root = tk.Tk()

    # hide the root window
    root.withdraw()

    # select a file using filedialog without a window
    file_path = filedialog.askopenfilename(filetypes=[('Choose file', '.xlsx .xlsm .xlsb .xltx .xltm .xls .xlt .xls .xml .xml .xlam .xla .xlw .xlr')])
    root.after(33, root.focus)

    # release memory
    root.destroy()
    root = None

    # return the selected file path
    return file_path


@eel.expose
def choose_folder():
    root = tk.Tk()

    root.withdraw()

    folder_path = filedialog.askdirectory()
    root.after(33, root.focus)

    root.destroy()
    root = None

    return folder_path


@eel.expose
def get_sheets(file):
    xl = pandas.ExcelFile(file)
    return xl.sheet_names


@eel.expose
def start(setup):
    process(setup)
    return 'Виконано'


some_path = os.path.dirname(os.path.realpath(__file__)).split('/')
some_path.remove(some_path[-1])
some_path = '/'.join(some_path)

eel.start(f'{some_path}/home.html', size=(500, 300))

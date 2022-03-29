from ast import Global, keyword
import tkinter as tk
import tkinter.filedialog as fd
from tkinter.tix import COLUMN

import openpyxl
import pydicom as pydicom
import pandas as pd
import os

win = tk.Tk()
# photo = tk.PhotoImage(file=r'D:\Dicom\interface\dicom.png')
# win.iconphoto(False, photo)
win.config(bg='#e6f2ff')

win.title('dicom tags')
win.geometry('1000x700+20+20')

#___________________________________________________________________
''' 
ФУНКЦИИ ДЛЯ ИНТЕРФЕЙСА
'''
def choose_dir_directory():
    global patient_dir
    directory = fd.askdirectory(title='Открыть папку', initialdir='\\').replace('/', '\\')
    if directory:
        #print(directory)
        patient_dir = directory
        tk.Label(win, text=f'Путь к папке: {directory}', font=20).grid(row=0, column=1, columnspan=2, stick='we')

def print_data():
    global tags, keyword
    keyword, tags = int(list_.curselection()[0]), tags_.get()
    if keyword:
        tags = tags.split()
    if not keyword:
        tags = [(int(tag[1:5], 16), int(tag[6:-1], 16)) for tag in tags.split()]

    print(f'Путь к папке пациента: {patient_dir}')
    print(f'Тип ключа: {list_.curselection()[0]}')
    print(f'Путь к выходной папаке: {excel_dir}')
    print(f'Теги: {tags_.get()}')

def choose_out_directory():
    global excel_dir
    directory = fd.askdirectory(title='Открыть папку', initialdir='\\').replace('/', '\\')
    if directory:
        excel_dir = directory + '\\' + file_name.get() + '.xlsx'
        #print(excel_dir)
        tk.Label(win, text=f'Путь к выходному файлу: {excel_dir}', font=20).grid(row=2, column=0, columnspan=3, stick='we')


#________________________________________________________
'''
Функции для консоли
'''

def parse_dicom_1(directory, tags, keyword):
    file_names = os.listdir(directory)
    direction_list = [directory + '\\' + file_name for file_name in os.listdir(directory)]
    
    if not keyword:
        tags = [(0x0008, 0x103e)] + tags
        cols = [pydicom.datadict.dictionary_keyword(tag) + f' ({tag[0]:04x}, {tag[1]:04x})' for tag in tags]
    
    if keyword:
        tags = ['SeriesDescription'] + tags
        temp_tags = [pydicom.tag.BaseTag(pydicom.datadict.tag_for_keyword(tag)) for tag in tags]
        cols = [pydicom.datadict.dictionary_keyword(temp_tags[i]) + ' ' +\
                str(pydicom.tag.BaseTag(pydicom.datadict.tag_for_keyword(tags[i]))) for i in range(len(tags))]
        
    df = pd.DataFrame(index=file_names, columns=cols)
    for i in range(len(file_names)):
        for j in range(len(tags)):
            df.iloc[i, j] = get_tag_value(direction_list[i], tags[j])
            
    return df


def get_tag_value(dir_, tag):
    dcm_file = pydicom.dcmread(dir_)
    try: return str(dcm_file[tag].value) if type(dcm_file[tag].value) == pydicom.valuerep.PersonName else dcm_file[tag].value
    except Exception:
        return '?'


def parse_by_patients(directory, tags, keyword):
    patients = os.listdir(directory)
    directories = [directory + '\\' + way for way in patients]
    dataframes = []
    dirs = []
    
    for dir_ in directories:
        df = (parse_dicom_1(dir_, tags, keyword))
        df = df.groupby('SeriesDescription (0008, 103e)').first().reset_index()
        df.insert(0, 'Patient', dir_.split('\\')[-1])
        df.loc[df.shape[0]] = ['-'] * df.shape[1]
        dataframes.append(df)
        
    return dataframes


def create_all_patients_df(dataframes):
    return pd.concat(dataframes)


def parse_by_patients_to_exel(patient_dir, excel_dir, tags, keyword):
    dataframes = parse_by_patients(patient_dir, tags, keyword)
    df = create_all_patients_df(dataframes)
    df.to_excel(excel_dir, index=False)
    return df

#___________________________________________________________________

def f():
    parse_by_patients_to_exel(patient_dir, excel_dir, tags, keyword)
    tk.Label(win, text='Готово', bg='green', font=20).grid(row=5, column=2)

#___________________________________________________________________

btn_dir = tk.Button(win, text='Выбрать папку для исследования', command=choose_dir_directory, font=20)
btn_dir.grid(row=0, column=0, stick='we')

dir_label = tk.Label(win, text='Путь к папке: -', font=20)
dir_label.grid(row=0, column=1, columnspan=2, stick='wens')


btn_out = tk.Button(win, text='Выбрать папку сохранения файла', command=choose_out_directory, font=20)
btn_out.grid(row=1, column=2, stick='we')

tk.Label(win, text='Имя файла: ', font=20).grid(row=1, column=0, stick='wens')


file_name = tk.Entry(win, font=20)
file_name.grid(row=1, column=1, stick='wens')

tk.Label(win, text='Путь к выходному файлу: -', font=20).grid(row=2, column=0, columnspan=3, stick='we')


key_label = tk.Label(win, text='Как будут введены теги', font=20)
key_label.grid(row=3, column=0, stick='wens')

keys = ['Кодами', 'Словами']
list_ = tk.Listbox(win, height=2, selectmode=tk.SINGLE, font=20)
for i in keys:
    list_.insert(tk.END, i)
list_.grid(row=3, column=1, columnspan=2, stick='we')


tk.Label(win, text='Введите теги через пробел', font=20).grid(row=4, column=0, stick='we')

tags_ = tk.Entry(win, font=20)
tags_.grid(row=4, column=1, columnspan=2, stick='we')



win.grid_columnconfigure(0, minsize=300)
win.grid_columnconfigure(3, minsize=400)







tk.Button(win, text='Обновить вводимые данные \nи вывести их в консоль', command=print_data, font=20).grid(row=5, column=0)
tk.Button(win, text='результат', command=f, font=20).grid(row=5, column=1)



win.mainloop()

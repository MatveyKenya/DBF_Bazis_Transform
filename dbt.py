#---------------------------------------------------
# 27/04/2020 - начато
# программа по преобразованию файлов .dbf
# из формата БАЗИС-8 и БАЗИС-2021 - в - БАЗИС-7
#
# Версия 1.1 от 28/07/2021 --- исправил под DBF Базис 2021
#---------------------------------------------------
import dbf
import os
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
from tkinter import messagebox


class Main_Win(tk.Tk):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.dir_path = '' #путь к папке с дбф        
        self.field_names_v8 = ['code', 'name', 'ediz', 'calckol', 'coef', 'kolprod',
        'kolorder', 'price', 'priceprod', 'priceorder', 'note']
        self.field_names_v2021 = ['code', 'name', 'ediz', 'calckol', 'coef', 'priceprod', 'price', 'kolprod',
        'kolorder', 'priceorder', 'ediz_model', 'note']

    def init_ui(self):
        self.geometry('400x170+300+300')
        self.title('.DBF Bazis Transformator v1.0')
        self.resizable(0,0)
        if os.path.isfile('dbt.ico'): self.iconbitmap('dbt.ico')    
        btn_open = ttk.Button(self, text='Открыть папку с dbf', command=self.press_open_dir, width=40)
        btn_open.pack(pady=5)
        lbl_path = ttk.Label(self, text='Путь до папки с файлам *.dbf (Базис v8)')
        lbl_path.pack(pady=5)
        self.entry_path = ttk.Entry(self)
        self.entry_path.pack(fill=tk.X, padx=5, pady=5)        
        btn_start_transform = ttk.Button(self, text='Преобразовать dbf из v8 в v7', command=self.start_transform, width=40)
        btn_start_transform.pack(padx=5, pady=5)
        menu = tk.Menu()
        menu.add_command(label='Справка для новичка ёу', command=self.spravka)
        menu.add_separator()
        menu.add_command(label='о проге', command=self.about)
        self.configure(menu=menu)

    def spravka(self):
        text_spravka = 'Программа преобразует файлы *.dbf выгруженные из Базиса-8 '\
                        +'в файлы *.dbf структуры выгрузки из под Базис-7 '\
                        +'для дальнейшей загрузки в программу OZMP\n\n'\
                        +'Порядок работы:\n'\
                        +'1) выбрать папку с файлами dbf формата Базис-8\n'\
                        +'   ---1. через кнопку "Открыть папку с dbf" '\
                        +'(Важно! сами файлы там не показываются! нужно просто выбрать саму папку)\n'\
                        +'   ---2. скопировать путь в окне открытой папки Windows '\
                        +'и вставить в строку программы\n'\
                        +'2) нажать кнопку "Преобразовать dbf из v8 в v7"\n'\
                        +'3) в случае успеха появится окно Отчета о преобразовании\n'\
                        +'4) автоматически создается папка "_copy_dbf_bazis8" с копией старых файлов v.8'\
                        +'\n\n ЗА РАБОТУ, ТОВАРИСЧ !!!'                        
        messagebox.showinfo('Инструкция по программе', text_spravka)

    def about(self):
        messagebox.showinfo('DBT v1.0', 'версия 1.0 от 1.05.2020\n\
            matveykenya@gmail.com')
    
    def press_open_dir(self):    
        dir_path = filedialog.askdirectory()    
        if dir_path !='':
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, dir_path)
    
    def start_transform(self):        
        if os.path.isdir(self.entry_path.get()): #проверка коректности пути
            self.dir_path = self.entry_path.get() 
            #находим все файлы и определяем из них dbf8
            items = os.listdir(self.dir_path)
            all_dbf = list()
            for i in filter(lambda x: x.endswith('.dbf'), items): all_dbf.append(i)                             
            if all_dbf: #если есть хоть 1 дбф в папке проверяем что это дбф7 и действуем
                v8_dbf = list() # здесь будут имена файлов и УЖЕ преобразованные в Базис7 таблицы
                for i in all_dbf:
                    file_input_b8 = self.dir_path +'/'+ i #файл dbf выгрузка из Базис8
                    db8=dbf.Table(file_input_b8, codepage='cp866')
                    db8.open()
                    # -проверка структуры, что это именно dbf Базис_8 по списку имен полей                    
                    if db8.field_names==self.field_names_v8 or db8.field_names==self.field_names_v2021:
                        table = list()
                        for row in db8: #заполняем таблицу в формате Базис7 (без смысла ставим 0)
                            table.append(
                                (row['code'], row['name'], row['ediz'],
                                str(row['calckol']), '0', '0', row['coef'],
                                str(row['kolprod']), '0', '0')
                            )          
                        db8.close()
                        v8_dbf.append((i, table))
                if v8_dbf:                                       
                    dir_path_copy_dbf_v8 = self.dir_path + '/_copy_dbf_bazis8'
                    if not os.path.isdir(dir_path_copy_dbf_v8): os.mkdir(dir_path_copy_dbf_v8)
                    #переносим все файлы v8 в dir_path_copy_dbf_v8 (или удаляем если не нужны будут потом)
                    text_info = ''
                    count_of_files = 0
                    for i in v8_dbf:
                        old_file_name = self.dir_path +'/'+ i[0]
                        copy_file_name = dir_path_copy_dbf_v8 +'/'+ i[0]
                        os.replace(old_file_name, copy_file_name)
                        self.create_dbf_b7(old_file_name, i[1]) #передаем имя файла и таблицу
                        text_info += i[0] +'\n'
                        count_of_files += 1
                    text_info = 'Файлы dbf формата Базис8 в количестве '\
                                + str(count_of_files) +' шт:\n\n' + text_info \
                                + '\nуспешно преобразованы в формат Базис7'
                    messagebox.showinfo(self.dir_path, text_info)                    

                else: messagebox.showerror("Ошибка", "В этой Папке нет .dbf от Базис8 или Базис2021")
            else: messagebox.showerror("Ошибка", "В этой Папке нет ни одного Файла .dbf")                    
        else: messagebox.showerror("Ошибка", "Неверный Путь или Папка")

    def create_dbf_b7(self, file_output='', table=list()):
        if table:
            db = dbf.Table(file_output,
                '''kod C(10); name C(250); ediz C(10); awtokol C(10); soputkol C(10);
                handkol C(10); kf C(10); allkol C(10); price C(10); summa C(10)''',
                codepage='cp1251', dbf_type='vfp'
                ) # была кодировка cp1251, cp866
            db.open(mode=dbf.READ_WRITE)
            for datum in table: db.append(datum)
            db.close()
        else: print('table dbf='+file_output+' is empty')    


# --------__Запуск__-----------
if __name__=='__main__':
    app = Main_Win()
    app.mainloop()

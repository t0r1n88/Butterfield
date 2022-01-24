from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from docxtpl import DocxTemplate
from tkinter import ttk
import pandas as pd
import os
import re


class Window(Tk):
    """
    Основной класс для работы
    """
    fio = ''

    def __init__(self):
        super().__init__()
        self.title('Butterfield')
        self.geometry('600x800')
        self.resizable(False,False)

        # Создаем объект вкладок

        self.tab_control = ttk.Notebook(self)

        # Создаем вкладку справки о выплатах студентам
        self.tab_payment = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab_payment, text='Создание единичной справки о выплатах')
        self.tab_control.pack(expand=1, fill='both')

        # Добавляем виджеты на вкладку
        # Создаем метку для описания назначения программы
        self.lbl_hello = Label(self.tab_payment, text='Скрипт для создания справки о выплатах')
        self.lbl_hello.grid(column=0, row=0, padx=10, pady=25)

        # Создаем кнопку Выбрать шаблон

        self.btn_template_contract = Button(self.tab_payment, text='1) Выберите шаблон справки',
                                            font=('Arial Bold', 20),
                                            command=self.select_file_template_payment_student
                                            )
        self.btn_template_contract.grid(column=0, row=1, padx=10, pady=10)

        # Создаем кнопку Выбрать файлы с данными
        self.btn_data_contract = Button(self.tab_payment, text='2) Выберите файлы с данными', font=('Arial Bold', 20),
                                        command=self.select_file_data_payment
                                        )
        self.btn_data_contract.grid(column=0, row=2, padx=10, pady=10)

        # Создаем кнопку для выбора папки куда будут генерироваться справка

        self.btn_choose_end_folder_contract = Button(self.tab_payment, text='3) Выберите конечную папку',
                                                     font=('Arial Bold', 20),
                                                     command=self.select_end_folder_payment_student
                                                     )
        self.btn_choose_end_folder_contract.grid(column=0, row=3, padx=10, pady=10)

        # Создаем поле для ввода ФИО
        # Текст который будет по умолчанию в текстовом поле
        self.input_field = Entry(self.tab_payment, justify='center', width=40, font=20)
        self.input_field.grid(column=0, row=4, padx=10, pady=10)
        self.input_field.focus()

        # # Создаем логотип
        # self.canvas = Canvas(tab_payment, height=320, width=550)
        # img = PhotoImage(file='logo.png')
        # self.image = self.canvas.create_image(0, 0, anchor='nw', image=img)
        # self.canvas.grid(column=1, row=0)

        # Создаем кнопку для запуска функции генерации файлов

        self.btn_create_file_payment = Button(self.tab_payment, text='4) Создать справку', font=('Arial Bold', 20),
                                              command=self.excecute_script)

        self.btn_create_file_payment.grid(column=0, row=5, padx=10, pady=10)

        # Создаем вкладку для массового создания справок
        self.tab_payment_group = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab_payment_group, text='Массовое создание справок о выплатах')
        self.tab_control.pack(expand=1, fill='both')

        # Добавляем виджеты на вкладку
        # Создаем метку для описания назначения программы
        self.lbl_hello = Label(self.tab_payment_group, text='Скрипт для массового создания справок о выплатах')
        self.lbl_hello.grid(column=0, row=0, padx=10, pady=25)

        # Создаем кнопку Выбрать шаблон

        self.btn_group_template_contract = Button(self.tab_payment_group, text='1) Выберите шаблон справки',
                                                  font=('Arial Bold', 20),
                                                  command=self.select_file_template_payment_student
                                                  )
        self.btn_group_template_contract.grid(column=0, row=1, padx=10, pady=10)

        # Создаем кнопку Выбрать файлы с данными
        self.btn_group_data_contract = Button(self.tab_payment_group, text='2) Выберите файлы с данными',
                                              font=('Arial Bold', 20),
                                              command=self.select_file_data_payment
                                              )
        self.btn_group_data_contract.grid(column=0, row=2, padx=10, pady=10)

        # Создаем кнопку для выбора папки куда будут генерироваться справка

        self.btn_choose_group_end_folder_contract = Button(self.tab_payment_group, text='3) Выберите конечную папку',
                                                           font=('Arial Bold', 20),
                                                           command=self.select_end_folder_payment_student
                                                           )
        self.btn_choose_group_end_folder_contract.grid(column=0, row=3, padx=10, pady=10)

        # Создаем кнопку для выбора файла со списком фамилий
        self.btn_choose_fio_lst = Button(self.tab_payment_group, text='4) Выберите файл со списком фамилий',
                                         font=('Arial Bold', 20),
                                         command=self.select_fio_lst
                                         )
        self.btn_choose_fio_lst.grid(column=0, row=4, padx=10, pady=10)

        # Создаем кнопку для создания справок
        # Создавать отдельную функцию отличающуюся только способом ввода ФИО такое себе конечно, в будущем может быть
        self.btn_group_create_file_payment = Button(self.tab_payment_group, text='5) Создать справки',
                                                    font=('Arial Bold', 20),
                                                    command=self.execute_script_group)

        self.btn_group_create_file_payment.grid(column=0, row=5, padx=10, pady=10)

    def execute_script_group(self):
        """
        Метод для отправки фио Student.generate_list_payment.
        Фио берутся из одноколоночного списка эксель
        :return:
        """
        # Считываем датафрейм и убираем пробельные символы
        df = pd.read_excel(program_window.fio_lst)
        df = df.applymap(lambda x: x.strip(), na_action='ignore')

        for row in df.itertuples():
            # row[1] ФИО
            # row[2] Группа
            Student.generate_list_payment(row[1], row[2])
        messagebox.showinfo('Итог операции', 'Справки созданы')

    def excecute_script(self):
        """
        Метод чтобы получиьт текст из поля ввода и отправить его статическому методу, ура костыли
        :return:
        """

        fio = self.input_field.get()
        Student.generate_list_payment(fio)

    @classmethod
    def select_fio_lst(cls):
        """
        Метод для выбора файла со списком фамилий студентов для которых нужно сделать справки
        :return: путь к файлу
        """
        cls.fio_lst = filedialog.askopenfilename(
            filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

    @classmethod
    def select_file_template_payment_student(cls):
        """
        Метод для выбора файла шаблона по которому будут генерироваться справки
        :return: Путь к файлу шаблона
        """
        cls.template_payment = filedialog.askopenfilename(
            filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

    @classmethod
    def select_file_data_payment(cls):
        """
        Метод для выбора файлов из которых будут генерироватся справки
        :return: Путь к файлам  с данными
        """
        cls.lst_data_files = filedialog.askopenfilenames(
            filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

    @classmethod
    def select_end_folder_payment_student(cls):
        """
        Метод для выбора папки куда будет генерироватся справка
        :return: путь к конечной папке
        """
        cls.end_folder = filedialog.askdirectory()


class Student:
    """
    Класс  ищущий конткретного студента и обрабатывающий его данные. На выходе будет список
    """

    @staticmethod
    def generate_list_payment(fio, group=''):
        """
  #          Метод для генерации справки студента
    :param fio : строка с ФИО студента
    :param group : необязательный параметр группа студента
  #         :return: список
  #         """
        print(fio)
        template = Window.template_payment
        lst_files = Window.lst_data_files
        end_folder = Window.end_folder
        # Список листов который будет доступен для всех экзмепляров класса Student
        lst_lists = ['3 корпус', 'гл корпус', '1 корпус', 'Хоринск', 'Федералы']
        # Содаем списки куда будут заносится результаты обработанные данные
        akadem_payment_lst = []
        social_payment_lst = []
        for file in lst_files:
            # находим год обрабатываемого файла
            year = re.search(r'\d{4}', file).group()
            # Начинаем поиск. Для этого в вложенном  цикле открываем листы каждого файла
            for i in range(len(lst_lists)):
                # С помощью skiprows пропускаем первую строку
                df = pd.read_excel(f'{file}', sheet_name=lst_lists[i], skiprows=1, dtype={'Ф.И.О.': 'str'})
                # заполняем пустые ячейки
                df = df.fillna(0.0)
                # Убираем ячейки с значением 0.0 в столбце ФИО, так как мы будем потом очищать их от пробелов и использовать эту колонку как индекс
                df = df[df['Ф.И.О.'] != 0.0]
                # Убираем пробельные символы в столбце  ФИО, чтобы потом не путаться
                df['Ф.И.О.'] = df['Ф.И.О.'].apply(lambda x: x.strip())
                # Осуществляем поиск
                result = Student.find_student(df, fio)
                # Если поиск ничего не нашел то ищем на следующем листе
                if result:
                    # Получаем кортеж состоящий из списка академических выплат за год и списка социальных выплат за год
                    year_payment = Student.processing_payment(result, year)
                    akadem_payment_lst.extend(year_payment[0])
                    social_payment_lst.extend(year_payment[1])
                    # print(akadem_payment_lst)
                    # print('**********')
                    # print(social_payment_lst)
                    # Выходим из списка ищущего в листах текущего файла и начинаем поиск в других файлах
                    break
                else:
                    continue
        # Считываем шаблон
        # print(akadem_payment_lst)
        # print('**********')
        # print(social_payment_lst)
        # Приводим к удобному виду
        output_lst_academ = [f'{month[0]} {month[1]} руб.' for month in akadem_payment_lst]
        output_lst_social = [f'{month[0]} {month[1]} руб.' for month in social_payment_lst]

        # Чтобы при массовом создании не спамить окна сообщений поставим условие что если group не равно '' то
        # выдаем сообщение только по окончании  генерации
        # Если мы получили результат то сохраняем справку. После окончания генерации показываем сообщение о успещном завершении
        # если студент не найден то показываем предупреждение.
        if group == '':
            if output_lst_academ and output_lst_social:
                doc = DocxTemplate(template)
                # заполняем словарь

                context = {'FIO': fio, 'GROUP': group, 'lst_payment_academ': output_lst_academ,
                           'lst_payment_social': output_lst_social}
                doc.render(context)
                doc.save(f'{end_folder}/{fio}.docx')
                messagebox.showinfo('Итог операции', 'Справка создана!')
            else:
                messagebox.showwarning('Итог операции', ' Студент не найден \n Проверьте правильность написания ФИО')
        else:
            if output_lst_academ and output_lst_social:
                doc = DocxTemplate(template)
                # заполняем словарь

                context = {'FIO': fio, 'GROUP': group, 'lst_payment_academ': output_lst_academ,
                           'lst_payment_social': output_lst_social}
                doc.render(context)
                doc.save(f'{end_folder}/{fio}.docx')
                print('Справка создана!')

    @staticmethod
    def find_student(df, fio):
        """
        Метод для поиска студента в датафрейме
        :param df: входной датафрейм
        :fio: ФИО студента
        :return: строка с данными студента если он найден, пустая строка если не найден
        """

        # Копируем датафрейм на всякий случай
        # Устанавливаем ФИО как индекс
        find_df = df.set_index('Ф.И.О.').copy()
        # Устанавливаем ФИО как индекс
        # find_df.to_excel('fio_index.xlsx')
        try:
            result_find = find_df.loc[fio]
            # Превращаем в словарь
            dct_result_find = result_find.to_dict()
            # Возвращаем найденую строку
            return dct_result_find

        except KeyError:
            # если такого индекса(ФИО) нет то возвращаем пустой словарь
            no_result_find = dict()
            return no_result_find

    @staticmethod
    def processing_payment(data: dict, year):
        """
        Функция для обработки словаря с данными студента
        :param data: словарь
        :param year: год обрабаотываемого файла
        :return: список? или словарь?
        """
        academ_temp_lst = Student.extract_payment(data, 'акад', 'прем', year)
        socical_temp_lst = Student.extract_payment(data, 'соц', 'сир', year)
        return (academ_temp_lst, socical_temp_lst)

    @staticmethod
    def extract_payment(data, word, word2, year):
        """
        Функция для первичной обработки
        :param data:
        :param word: ключевое слово типа стипендии(акад,соц)
        :param word2: дополнительное ключевое слово(сир или прем)
        :param year: год обрабатываемого файла
        :return:Список кортежей формата месяц год,сумм выплат
        """
        # Список куда будут заносится вычисленные данные
        output_lst = []

        month_dict = dict()
        # заполнили словарь с названиями месяцев и годом
        year_digit = int(year)
        month_dict = {f'январь {year_digit}г': 0, f'февраль {year_digit}г': 0, f'март {year_digit}г': 0, f'апрель {year_digit}г': 0, f'май {year_digit}г': 0,
                      f'июнь {year_digit}г': 0, f'июль {year_digit}г': 0,
                      f'август {year_digit}г': 0, f'сентябрь {year_digit}г': 0, f'октябрь {year_digit}г': 0, f'ноябрь {year_digit}г': 0,
                      f'декабрь {year_digit}г': 0, }



        # if int(year) == 2019:
        #     month_dict = {'январь 2019г': 0, 'февраль 2019г': 0, 'март 2019г': 0, 'апрель 2019г': 0, 'май 2019г': 0,
        #                   'июнь 2019г': 0, 'июль 2019г': 0,
        #                   'август 2019г': 0, 'сентябрь 2019г': 0, 'октябрь 2019г': 0, 'ноябрь 2019г': 0,
        #                   'декабрь 2019г': 0, }
        # elif int(year) == 2020:
        #     month_dict = {'январь 2020г': 0, 'февраль 2020г': 0, 'март 2020г': 0, 'апрель 2020г': 0, 'май 2020г': 0,
        #                   'июнь 2020г': 0, 'июль 2020г': 0,
        #                   'август 2020г': 0, 'сентябрь 2020г': 0, 'октябрь 2020г': 0, 'ноябрь 2020г': 0,
        #                   'декабрь 2020г': 0, }
        # elif int(year) == 2021:
        #     month_dict = {'январь 2021г': 0, 'февраль 2021г': 0, 'март 2021г': 0, 'апрель 2021г': 0, 'май 2021г': 0,
        #                   'июнь 2021г': 0, 'июль 2021г': 0,
        #                   'август 2021г': 0, 'сентябрь 2021г': 0, 'октябрь 2021г': 0, 'ноябрь 2021г': 0,
        #                   'декабрь 2021г': 0, }
        # elif int(year) == 2022:
        #     month_dict = {'январь 2022г': 0, 'февраль 2022г': 0, 'март 2022г': 0, 'апрель 2022г': 0, 'май 2022г': 0,
        #                   'июнь 2022г': 0, 'июль 2022г': 0,
        #                   'август 2022г': 0, 'сентябрь 2022г': 0, 'октябрь 2022г': 0, 'ноябрь 2022г': 0,
        #                   'декабрь 2022г': 0, }
        # elif int(year) == 2023:
        #     month_dict = {'январь 2023г': 0, 'февраль 2023г': 0, 'март 2023г': 0, 'апрель 2023г': 0, 'май 2023г': 0,
        #                   'июнь 2023г': 0, 'июль 2023г': 0,
        #                   'август 2023г': 0, 'сентябрь 2023г': 0, 'октябрь 2023г': 0, 'ноябрь 2023г': 0,
        #                   'декабрь 2023г': 0, }
        # elif int(year) == 2024:
        #     month_dict = {'январь 2024г': 0, 'февраль 2024г': 0, 'март 2024г': 0, 'апрель 2024г': 0, 'май 2024г': 0,
        #                   'июнь 2024г': 0, 'июль 2024г': 0,
        #                   'август 2024г': 0, 'сентябрь 2024г': 0, 'октябрь 2024г': 0, 'ноябрь 2024г': 0,
        #                   'декабрь 2024г': 0, }
        # elif int(year) == 2025:
        #     month_dict = {'январь 2025г': 0, 'февраль 2025г': 0, 'март 2025г': 0, 'апрель 2025г': 0, 'май 2025г': 0,
        #                   'июнь 2025г': 0, 'июль 2025г': 0,
        #                   'август 2025г': 0, 'сентябрь 2025г': 0, 'октябрь 2025г': 0, 'ноябрь 2025г': 0,
        #                   'декабрь 2025г': 0, }
        # elif int(year) == 2026:
        #     month_dict = {'январь 2026г': 0, 'февраль 2026г': 0, 'март 2026г': 0, 'апрель 2026г': 0, 'май 2026г': 0,
        #                   'июнь 2026г': 0, 'июль 2026г': 0,
        #                   'август 2026г': 0, 'сентябрь 2026г': 0, 'октябрь 2026г': 0, 'ноябрь 2026г': 0,
        #                   'декабрь 2026г': 0, }

        # Заполняем
        for key_month, value_month in month_dict.items():
            for key_data, value_data in data.items():
                # если ключ месяца(например май 2019) есть в ключе базового словаря то значение этого ключа суммируем
                # try на случай некорректных данных
                if (key_month in key_data) and (word in key_data or word2 in key_data):
                    try:
                        # print(float(data[key_data]))
                        month_dict[key_month] += float(data[key_data])

                    except:
                        print(key_data, value_data)
                        continue
        # print(month_dict)

        # Конвертируем в список, чтобы сохранять последовательность.
        temp_lst = list(month_dict.items())
        return temp_lst


if __name__ == '__main__':
    program_window = Window()

    program_window.mainloop()

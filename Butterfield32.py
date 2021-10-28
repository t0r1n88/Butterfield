from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from docxtpl import DocxTemplate
from tkinter import ttk
import pandas as pd


def select_file_template_payment_student():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_payment_student
    name_file_template_contracts = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_payment():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться договор  и генерация падежей фио
    :return: Путь к файлу с данными и словарь с просклоняемыми ФИО
    """
    global name_file_data_payment
    # Получаем путь к файлу
    name_file_data_payment = filedialog.askopenfilenames(parent =window,filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
    print(name_file_data_payment)

def select_end_folder_payment_student():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_payment_student
    path_to_end_folder_payment_student = filedialog.askdirectory()


def generate_list_payment():
    """
    Функция для генерации справки студента
    :return:
    """
    print(input_field.get())

if __name__ == '__main__':
    window = Tk()
    window.title('Butterfield')
    window.geometry('1024x768')


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку справки о выплатах студентам
    tab_payment = ttk.Frame(tab_control)
    tab_control.add(tab_payment, text='Создание справок о выплатах')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_payment, text='Скрипт для создания справок о выплатах')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон

    btn_template_contract = Button(tab_payment, text='1) Выберите шаблон справки', font=('Arial Bold', 20),
                                   command=select_file_template_payment_student
                                   )
    btn_template_contract.grid(column=0, row=1, padx=10, pady=10)


    # Создаем кнопку Выбрать файлы с данными
    btn_data_contract = Button(tab_payment, text='2) Выберите файлы с данными', font=('Arial Bold', 20),
                               command=select_file_data_payment
                               )
    btn_data_contract.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться справка

    btn_choose_end_folder_contract = Button(tab_payment, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                            command=select_end_folder_payment_student
                                            )
    btn_choose_end_folder_contract.grid(column=0, row=3, padx=10, pady=10)


    # Создаем поле для ввода ФИО
    # Текст который будет по умолчанию в текстовом поле
    # default_message = StringVar()
    # default_message.set('Введите ФИО студента')
    input_field = Entry(tab_payment, justify='center', width=40, font=20)
    input_field.grid(column=0, row=4, padx=10, pady=10)



    # Создаем логотип
    canvas = Canvas(tab_payment,height=320,width=550)
    img = PhotoImage(file='logo.png')
    image = canvas.create_image(0,0,anchor='nw',image=img)
    canvas.grid(column=1,row=0)

    # Создаем кнопку для запуска функции генерации файлов

    btn_create_file_payment = Button(tab_payment, text='4) Создать справку', font=('Arial Bold', 20),
                                       command=generate_list_payment)
    btn_create_file_payment.grid(column=0, row=5, padx=10, pady=10)
    #
    # # Создаем вкладку для создания приказов о зачислении
    # tab_order_enroll = ttk.Frame(tab_control)
    # tab_control.add(tab_order_enroll, text='Создание приказов о зачислении')
    #
    # # Добавляем виджеты на вкладку
    # lbl_hello = Label(tab_order_enroll, text='Скрипт для создания приказов о зачислении')
    # lbl_hello.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Создаем кнопку Выбрать шаблон
    #
    # btn_template_scc = Button(tab_order_enroll, text='1) Выберите шаблон приказа', font=('Arial Bold', 20),
    #                           command=select_file_template_order_enroll, )
    # btn_template_scc.grid(column=0, row=1, padx=10, pady=10)
    #
    # # Создаем кнопку Выбрать файл с данными
    # btn_data_scc = Button(tab_order_enroll, text='2) Выберите файл с данными', font=('Arial Bold', 20),
    #                       command=select_file_data_order_enroll)
    # btn_data_scc.grid(column=0, row=2, padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_folder_scc = Button(tab_order_enroll, text='3) Выберите конечную папку', font=('Arial Bold', 20),
    #                                    command=select_end_folder_order_enroll)
    # btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Создаем кнопку для запуска функции генерации файлов
    #
    # btn_create_files_scc = Button(tab_order_enroll, text=' Создать приказы', font=('Arial Bold', 20),
    #                               command=generate_order_enroll)
    # btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

    #
    #
    #

    window.mainloop()

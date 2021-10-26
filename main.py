import pandas as pd
import os
import re


def extract_payment(data, word, year):
    """
    Функция для первичной обработки
    :param data:
    :param word: ключевое слово типа стипендии(акад,соц,сирот)
    :param year: год обрабатываемого файла
    :return:
    """
    # Список куда будут заносится вычисленные данные
    output_lst = []
    # Так как будут использоваться несколько файлов придется использовать регулярки
    month_dict =dict()

    for key,value in data.items():
        result =

def processing_payment(data: dict, year):
    """
    Функция для обработки словаря с данными студента
    :param data: словарь
    :param year: год обрабаотываемого файла
    :return: список? или словарь?
    """
    academ_temp_lst = extract_payment(data, 'акад',year)


def find_student(df, fio):
    """
    Функция для поиска студента в датафрейме
    :param df: входной датафрейм
    :fio: ФИО студента
    :return: строка с данными студента если он найден, пустая строка если не найден
    """
    # Создаем пустой словарь
    # result_find = dict()
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
        # если такого индекса нет то возвращаем пустой словарь
        no_result_find = dict()
        return no_result_find


# Считываем файлы
# Путь к файлам
path = 'resources/'
# Тестовый студент
# fio = 'Зоз-Ткачук Алексей Валентинович'
fio = 'Будаев Борис Дармаевич '
# Список обрабатываемых листов
# lst_lists = ['3 корпус', 'гл корпус', '1 корпус', '1 корпус', '1 корпус']
lst_lists = ['3 корпус']

for file in os.listdir(path):
    # Содаем списки куда будуту заносится результаты обработанные данные
    akadem_payment_lst = []
    social_payment_lst = []
    orphan_payment_lst = []
    # находим год обрабатываемого файла
    year = re.search(r'\d{4}', file).group()
    # Начинаем поиск. Для этого в вложенном  цикле открываем листы каждого файла
    for i in range(len(lst_lists)):
        # С помощью skiprows пропускаем первую строку
        df = pd.read_excel(f'{path}{file}', sheet_name=lst_lists[i], skiprows=1)
        # Осуществляем поиск
        result = find_student(df, fio)
        print(result)
        # Если поиск ничего не нашел то ищем на следующем листе
        if result:
            # Получаем список
            year_payment = processing_payment(result, year)
        else:
            continue

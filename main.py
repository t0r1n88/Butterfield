import pandas as pd
import os
import re


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
    if int(year) == 2019:
        month_dict = {'январь 2019г': 0, 'февраль 2019г': 0, 'март 2019г': 0, 'апрель 2019г': 0, 'май 2019г': 0,
                      'июнь 2019г': 0, 'июль 2019г': 0,
                      'август 2019г': 0, 'сентябрь 2019г': 0, 'октябрь 2019г': 0, 'ноябрь 2019г': 0,
                      'декабрь 2019г': 0, }
    elif int(year) == 2020:
        month_dict = {'январь 2020г': 0, 'февраль 2020г': 0, 'март 2020г': 0, 'апрель 2020г': 0, 'май 2020г': 0,
                      'июнь 2020г': 0, 'июль 2020г': 0,
                      'август 2020г': 0, 'сентябрь 2020г': 0, 'октябрь 2020г': 0, 'ноябрь 2020г': 0,
                      'декабрь 2020г': 0, }
    elif int(year) == 2021:
        month_dict = {'январь 2021г': 0, 'февраль 2021г': 0, 'март 2021г': 0, 'апрель 2021г': 0, 'май 2021г': 0,
                      'июнь 2021г': 0, 'июль 2021г': 0,
                      'август 2021г': 0, 'сентябрь 2021г': 0, 'октябрь 2021г': 0, 'ноябрь 2021г': 0,
                      'декабрь 2021г': 0, }

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


def processing_payment(data: dict, year):
    """
    Функция для обработки словаря с данными студента
    :param data: словарь
    :param year: год обрабаотываемого файла
    :return: список? или словарь?
    """
    academ_temp_lst = extract_payment(data, 'акад', 'прем', year)
    socical_temp_lst = extract_payment(data, 'соц', 'сир', year)
    return (academ_temp_lst, socical_temp_lst)


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
        # если такого индекса(ФИО) нет то возвращаем пустой словарь
        no_result_find = dict()
        return no_result_find


# Считываем файлы
# Путь к файлам
path = 'resources/'
# Тестовый студент
# fio = 'Зоз-Ткачук Алексей Валентинович'
# 3 корпус
# fio = 'Будаев Борис Дармаевич '
# гл.корпус
# fio = 'Коротков Роман Сергеевич'
# 1 корпус
# fio = 'Зангеев Дмитрий Николаевич'
# Хоринск
# fio = 'Будаев Борис Дармаевич '
# Федералы
fio = 'Минеева Дарья Ивановна'
# Список обрабатываемых листов
# lst_lists = ['3 корпус', 'гл корпус', '1 корпус', 'Хоринск', 'Федералы']
lst_lists = ['3 корпус', 'гл корпус', '1 корпус', 'Хоринск', 'Федералы']
# Содаем списки куда будут заносится результаты обработанные данные
akadem_payment_lst = []
social_payment_lst = []
for file in os.listdir(path):
    # находим год обрабатываемого файла
    year = re.search(r'\d{4}', file).group()
    print(file)
    # Начинаем поиск. Для этого в вложенном  цикле открываем листы каждого файла
    for i in range(len(lst_lists)):
        # С помощью skiprows пропускаем первую строку
        df = pd.read_excel(f'{path}{file}', sheet_name=lst_lists[i], skiprows=1, dtype={'Ф.И.О.': 'str'})
        # заполняем пустые ячейки
        print(lst_lists[i])
        df = df.fillna(0.0)
        # Убираем ячейки с значением 0.0 в столбце ФИО, так как мы будем потом очищать их от пробелов и использовать эту колонку как индекс
        df = df[df['Ф.И.О.'] != 0.0]
        # Убираем пробельные символы в столбце  ФИО, чтобы потом не путаться
        df['Ф.И.О.'] = df['Ф.И.О.'].apply(lambda x: x.strip())
        # Осуществляем поиск
        result = find_student(df, fio)
        # Если поиск ничего не нашел то ищем на следующем листе
        if result:
            print(result)
            # Получаем кортеж состоящий из списка академических выплат за год и списка социальных выплат за год
            year_payment = processing_payment(result, year)
            akadem_payment_lst.extend(year_payment[0])
            social_payment_lst.extend(year_payment[1])
            print(akadem_payment_lst)
            print(social_payment_lst)
            # Выходим из списка ищущего в листах текущего файла и начинаем поиск в других файлах
            break

        else:
            continue

import requests as req
import pandas as pd
import matplotlib.pyplot as plt
import json
from datetime import datetime
from tokens import *


def take_start_date(attempt: int = 0):
    """
    :return: text start date
    """
    attempt += 1
    now = datetime.now()
    start_date_text = input('Введите дату в формате дд.мм.гггг (по умолчанию первое число текущего месяца): ')
    if start_date_text == '':
        start_date_text = f'01.{now.month}.{now.year}'
    try:
        text_date = datetime.strptime(start_date_text, '%d.%m.%Y')
    except ValueError:
        print(f'You must insert date in format dd.mm.YYYY, {5 - attempt} attempt left')
        if attempt > 4:
            print('attempts ended the program closes')
            exit()
        else:
            start_date_text = take_start_date(attempt)
    if text_date < datetime.strptime('01.01.2015', '%d.%m.%Y') or text_date > now:
        print(f'Вы ввели не корректную дату ({start_date_text}), будет назначена дата по умолчанию')
        start_date_text = f'01.{now.month}.{now.year}'
    return start_date_text


def get_data(filter_date: str):
    """
    :param filter_date: start date filter
    :return: dictionary
    """
    # Данные для запроса и авторизации
    auth = (email, api_key)
    url = f'https://mycompany.omnidesk.ru/api/cases.json?from_time={filter_date}'
    headers = {'Content-Type': 'application/json'}

    # Отправляем запрос
    omni = req.get(url=url, headers=headers, auth=auth)
    response = omni.text
    return json.loads(response)


def clear_df(dataframe):
    """
    :param dataframe: dirty dataframe
    :return: clear dataframe without open applications, with integer in closing speed and with integer in priority
    """
    # Убираем открытые заявки они для анализа в текущий момент не нужны
    close_only = dataframe[dataframe['closing_speed'] != '-']

    # Преобразуем в числовое значение столбец с минутами
    close_only['closing_speed'] = close_only['closing_speed'].astype('int')
    """
    Для решения вопроса сортировки по важности необходимо каждую категорию важности пронумеровать
    для этого создадим словарь самое важное - самая большая цифра
    """
    dict_priority = {'low': 1, 'normal': 2, 'high': 3, 'critical': 4}

    # все значения важности заменим на соответствующие значения из словаря
    close_only['priority'] = close_only['priority'].apply(lambda x: dict_priority.get(x))
    return close_only


def create_df(dictionary: dict):
    """
    :param dictionary: dictionary from json object
    :return: dataframe
    """
    dataframe = pd.DataFrame()
    for i in range(dictionary['total_count']):
        dict_new = dictionary[f'{i}']['case']
        dataframe = dataframe.append(dict_new, ignore_index=True)
    return dataframe


def create_hist(dataframe):
    """
    :param dataframe:
    :return: closing_speed histogram
    """
    dataframe['closing_speed'].hist(bins=10)
    plt.savefig('fig2.png')
    plt.show()


start_date = take_start_date()
json_obj = get_data(start_date)
df = create_df(json_obj)
try:
    df = clear_df(df)
    create_hist(df)
except KeyError:
    print('список заявок подходящих вашему запросу пуст')
    exit()
"""
Создадим новый датафрейм в котором будут только плохие заявки (время закрытия выше среднего)
Здесь я выбрал среднее так как у меня было открыто мало заявок и в этом случае это допустимо
У Вас возможно есть какой-то норматив времени ответа на заявку, с определенными критериями
например по важности
"""
bad_appeal = df[df['closing_speed'] > df['closing_speed'].mean()]

# Отсортируем по двум столбцам важности и скорости закрытия вверху самые важные и долгие
bad_appeal = bad_appeal.sort_values(by=['priority', 'closing_speed'], ascending=False)

# Сохраним последний датафрейм в excel
bad_appeal.to_excel('bad_appeal.xls', sheet_name='bad_appeal', engine='openpyxl')

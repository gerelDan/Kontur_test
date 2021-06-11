import numpy as np

import requests as req

import json

import pandas as pd

import matplotlib.pyplot as plt

from datetime import datetime

from tokens import *


start_date = input('Введите дату (по умолчанию первое число текущего месяца): ')
if start_date == '':
    now = datetime.now()
    start_date = f'01.{now.month}.{now.year}'
#Данные для авторизации


auth = (email, api_key)
#Ссылка для запроса
url = f'https://mycompany.omnidesk.ru/api/cases.json?from_time={start_date}'
#Заголовок запроса
headers = {'Content-Type': 'application/json'}
#Отправляем запрос с нужными данными
omni = req.get(url=url, headers=headers, auth=auth)
#Получаем текст ответа
response = omni.text
#преобразуем текст ответа в словарь
json_obj = json.loads(response)

#задаем столбцы для датафрейма
columns = list(json_obj['0']['case'].keys())
data = []
#пробегаемся циклом по всем заявкам
for i in range(json_obj['total_count']):
    row = []
    i_str = str(i)
    #Добавляем в список значения всех столбцов
    for key in json_obj[i_str]['case']:
        row.append(json_obj[i_str]['case'][key])
    #Делаем список списков для датафрейма
    data.append(row)

#создаем датафрейм на основе данных
df = pd.DataFrame(data=data, columns=columns)
#Убираем открытые заявки они для анализа в текущий момент не нужны
df_close_only = df[df['closing_speed'] != '-']
#Преобразуем в числовое значение столбец с минутами
df_close_only['closing_speed'] = df_close_only['closing_speed'].astype('int')
#Зададим гистограмму
df_close_only['closing_speed'].hist(bins=10)
#Сохраним гистограмму
plt.savefig('fig2.png')
#Покажем гистограмму
plt.show()
"""
Для решения вопроса сортировки по важности необходимо каждую категорию важности пронумеровать
для этого создадим словарь самое важное - самая большая цифра
"""
dict_priority = {'low': 1, 'normal': 2, 'high': 3, 'critical': 4}
#Добавим новый столбец под цифровой эквивалент важности
df_close_only['priority_num'] = ''
#пробежимся по всем вариантам важности и проставим в последний столбец соответствующие значения
for i in dict_priority:
    df_close_only['priority_num'] = np.where(df_close_only['priority'] == i,
                                             dict_priority[i], df_close_only['priority_num'])
#Узнаем среднее время закрытия заявки
mean_time = df_close_only['closing_speed'].mean()
"""
Создадим новый датафрейм в котором будут только плохие заявки (время закрытия выше среднего)
Здесь я выбрал среднее так как у меня было открыто мало заявок и в этом случае это допустимо
У Вас возможно есть какой-то норматив времени ответа на заявку, с определенными критериями
например по важности
"""
bad_appeal = df_close_only[df_close_only['closing_speed'] > mean_time]
#Отсортируем по двум столбцам важности и скорости закрытия вверху самые важные и долгие
bad_appeal = bad_appeal.sort_values(by=['priority_num', 'closing_speed'], ascending=False)
#Сохраним последний датафрейм в excel
bad_appeal.to_excel('bad_appeal.xls', sheet_name='bad_appeal', engine='openpyxl')

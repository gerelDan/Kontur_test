import pyodbc
from datetime import datetime
import pandas as pd

# Подключаемся к базе данных

# conn = pyodbc.connect(r'Driver={SQL Server};'
#                      r'Server=YourServer;' #Введите ваш сервер вместо YourServer
#                      r'Database=YourDataBase;' #Введите вашу базу вместо YourDatabase
#                      r'Trusted_Connection=yes;'
#                      )

def connect_data_base():
    return pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
                          r'DBQ=C:\Users\yermodan\PycharmProjects\Kontur\Database51.accdb;'
                          )

def get_table(query, cursor, filter_sql, now):
    table_from_sql = cursor.execute(query, filter_sql, now)
    data = []
    for i in table_from_sql.fetchall():
        data.append(list(i))
    desc = table_from_sql.description
    columns = []
    for row in desc:
        columns.append((row[0]))
    return data, columns


conn = connect_data_base()
cursor = conn.cursor()

# Пишем запрос на соединение всех таблиц
query_open_now = 'select num, bdate, pdate, cid, product, cost, payed, upto, tip from Bills LEFT JOIN (select ' \
            'Bill_content.bID, product, cost, payed, upto, tip from Bill_content LEFT JOIN retail_packs ON ' \
            '(retail_packs.bcID = Bill_content.bcID)) AS query_1 ON (query_1.bID = Bills.id)' \
            'where product = ? and upto is not NULL and upto >= ?'

query_close_now = 'select num, bdate, pdate, cid, product, cost, payed, upto, tip from Bills LEFT JOIN (select ' \
            'Bill_content.bID, product, cost, payed, upto, tip from Bill_content LEFT JOIN retail_packs ON ' \
            '(retail_packs.bcID = Bill_content.bcID)) AS query_1 ON (query_1.bID = Bills.id)' \
            'where product = ? and upto is not NULL and upto < ?'

# определяем время сейчас
now = datetime.now()
filter_sql = 'Контур-экстерн'

data_open_now = get_table(query_open_now, cursor, filter_sql, now)[0]
columns = get_table(query_open_now, cursor, filter_sql, now)[1]

data_close_now = get_table(query_close_now, cursor, filter_sql, now)[0]

# С оздадим датафреймы
df_open_now = pd.DataFrame(data=data_open_now, columns=columns)

df_close_now = pd.DataFrame(data=data_close_now, columns=columns)

# Проссумируем стомость поставок по счетам
df_open_now = df_open_now.groupby(
    ['num', 'bdate', 'pdate', 'cid', 'upto'], as_index=False).agg({'cost': 'sum', 'payed': 'sum'})
df_close_now = df_close_now.groupby(
    ['num', 'bdate', 'pdate', 'cid', 'upto'], as_index=False).agg({'cost': 'sum', 'payed': 'sum'})

# Оставим только такие счета где срок поставки максимальный
df_more_now_grouped = df_open_now.groupby(
    "cid", group_keys=False).apply(lambda x: x.nlargest(1, "upto"))
df_more_now_grouped = df_more_now_grouped[['cid', 'num', 'bdate', 'pdate', 'cost', 'payed', 'upto']]

# Оставим только такие счета где дата оплаты максимальная
df_less_now_grouped = df_close_now.groupby("cid", group_keys=False).apply(lambda x: x.nlargest(1, "pdate"))

# соединим две таблицы
df_all = pd.concat([df_more_now_grouped, df_less_now_grouped], axis=0, ignore_index=True)

# Еще раз оставим только такие счета где срок поставки максимальный
df_all = df_all.groupby("cid", group_keys=False).apply(lambda x: x.nlargest(1, "upto"))

# отсортируем по номерам клиентов
df_all = df_all.sort_values('cid', ignore_index=True)
df_all = df_all[['cid', 'num', 'bdate', 'pdate', 'cost', 'payed']]

# Получим столбцы из датафрейма
columns = df_all.columns

# Из датафрейма получим список списков
list_data = df_all.values.tolist()

# Закинем ексель результат
df_all.to_excel('result.xls', sheet_name='result', engine='openpyxl')

'''
Здесь результат отправим в базу данных
Попробуем удалить таблицу
'''
x = input('Пожалуйста удостоверьтесь, что в таблице с названием Result_kontur_ekstern нет критически важных данных'
          'если это так введите "Y" и нажмите enter:')
if x.lower() == 'y':
    try:
        cursor.execute('Drop table Result_kontur_ekstern')
        conn.commit()
    except pyodbc.ProgrammingError:
        pass
    except pyodbc.Error:
        print('таблица Result_kontur_ekstern открыта пользователемб закройте таблицу')
        exit()

# Создадим таблицу заново с нужными названиями и форматами
    create_query = f'create table Result_kontur_ekstern (' \
                   f'{columns[0]} int,' \
                   f' {columns[1]} int,' \
                   f' {columns[2]} datetime,' \
                   f' {columns[3]} datetime,' \
                   f' {columns[4]} money,' \
                   f' {columns[5]} money)'
    sql_open_now = cursor.execute(create_query)
    conn.commit()

# Преобразуем данные в подходящиц формат
    for i in range(len(list_data)):
        list_data[i][2] = datetime.strftime(datetime.strptime(str(list_data[i][2])[:10], '%Y-%m-%d'), '%Y-%m-%d')
        list_data[i][3] = datetime.strftime(datetime.strptime(str(list_data[i][3])[:10], '%Y-%m-%d'), '%Y-%m-%d')
        insert_execute = f'insert into Result_kontur_ekstern (' \
                         f'{columns[0]}, {columns[1]}, {columns[2]}, {columns[3]}, {columns[4]}, {columns[5]})' \
                         f'values ' \
                         f'({list_data[i][0]}, {list_data[i][1]}, {list_data[i][2]},' \
                         f' {list_data[i][3]}, {list_data[i][4]}, {list_data[i][5]})'
        insert_execute = str(insert_execute)
        cursor.execute(insert_execute)
        conn.commit()
    print('В базе данных проверьте таблицу Result_kontur_ekstern в ней результат работы скрипта')

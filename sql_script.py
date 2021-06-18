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


def get_table(query, cur, filtration, date_now):
    table_from_sql = cur.execute(query, filtration, date_now)
    data = []
    for row in table_from_sql.fetchall():
        data.append(list(row))
    desc = table_from_sql.description
    columns_table = []
    for row in desc:
        columns_table.append((row[0]))
    return data, columns_table


def work(df_open, df_close):
    # Проссумируем стомость поставок по счетам
    df_open = df_open.groupby(
        ['num', 'bdate', 'pdate', 'cid', 'upto'], as_index=False).agg({'cost': 'sum', 'payed': 'sum'})
    df_close = df_close.groupby(
        ['num', 'bdate', 'pdate', 'cid', 'upto'], as_index=False).agg({'cost': 'sum', 'payed': 'sum'})

    # Оставим только такие счета где срок поставки максимальный
    df_more_now_grouped = df_open.groupby(
        "cid", group_keys=False).apply(lambda cid: cid.nlargest(1, "upto"))
    df_more_now_grouped = df_more_now_grouped[['cid', 'num', 'bdate', 'pdate', 'cost', 'payed', 'upto']]

    # Оставим только такие счета где дата оплаты максимальная
    df_less_now_grouped = df_close.groupby("cid", group_keys=False).apply(lambda cid: cid.nlargest(1, "pdate"))

    # соединим две таблицы
    df_all = pd.concat([df_more_now_grouped, df_less_now_grouped], axis=0, ignore_index=True)

    # Еще раз оставим только такие счета где срок поставки максимальный
    df_all = df_all.groupby("cid", group_keys=False).apply(lambda cid: cid.nlargest(1, "upto"))

    # отсортируем по номерам клиентов
    df_all = df_all.sort_values('cid', ignore_index=True)
    df_all = df_all[['cid', 'num', 'bdate', 'pdate', 'cost', 'payed']]
    return df_all


def delete_table(cur, connect):
    x = input('Пожалуйста удостоверьтесь, что в таблице с названием Result_kontur_ekstern нет критически важных данных'
              'если это так введите "Y" и нажмите enter:')
    if x.lower() == 'y':
        try:
            cur.execute('Drop table Result_kontur_ekstern')
            connect.commit()
        except pyodbc.ProgrammingError:
            pass
        except pyodbc.Error:
            print('таблица Result_kontur_ekstern открыта пользователем закройте таблицу')
            exit()
    else:
        print('Вы не подтвердили отсутствие важных данных в таблице программа закрывается')
        exit()


def create_table(col, cur, connect):
    create_query = f'create table Result_kontur_ekstern (' \
                   f'{col[0]} int,' \
                   f' {col[1]} int,' \
                   f' {col[2]} datetime,' \
                   f' {col[3]} datetime,' \
                   f' {col[4]} money,' \
                   f' {col[5]} money)'
    cur.execute(create_query)
    connect.commit()


def fill_table (data, col, cur, connect):

# Преобразуем данные в подходящиц формат
    for i in range(len(data)):
        data[i][2] = str(data[i][2])[:10]
        data[i][3] = str(data[i][3])[:10]
        insert_execute = f'insert into Result_kontur_ekstern (' \
                         f'{col[0]}, {col[1]}, {col[2]}, {col[3]}, {col[4]}, {col[5]})' \
                         f'values ' \
                         f'({data[i][0]}, {data[i][1]}, {data[i][2]},' \
                         f' {data[i][3]}, {data[i][4]}, {data[i][5]})'
        insert_execute = str(insert_execute)
        cur.execute(insert_execute)
        connect.commit()



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

now = datetime.now()
filter_sql = 'Контур-экстерн'

data_open_now, columns = get_table(query_open_now, cursor, filter_sql, now)
data_close_now = get_table(query_close_now, cursor, filter_sql, now)[0]

df_open_now = pd.DataFrame(data=data_open_now, columns=columns)

df_close_now = pd.DataFrame(data=data_close_now, columns=columns)
df_finish = work(df_open_now, df_close_now)

columns = df_finish.columns

list_data = df_finish.values.tolist()

df_finish.to_excel('result.xls', sheet_name='result', engine='openpyxl')

delete_table(cursor, conn)

create_table(columns, cursor, conn)

fill_table(list_data, columns, cursor, conn)

print('В базе данных проверьте таблицу Result_kontur_ekstern в ней результат работы скрипта')

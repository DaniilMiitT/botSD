import json
import os

import pymssql
import redis


def try_convert_encoding(text, from_encoding='latin1', to_encoding='windows-1251'):
    """Пытается перекодировать строку. Если не удается — возвращает оригинал."""
    if isinstance(text, str):
        try:
            return text.encode(from_encoding).decode(to_encoding)
        except (UnicodeEncodeError, UnicodeDecodeError):
            
            return text
    return text


def convert_encoding(data, from_encoding='latin1', to_encoding='windows-1251'):
    """Recursively convert encoding for strings, with error handling."""
    if isinstance(data, str):
        return try_convert_encoding(data, from_encoding, to_encoding)
    elif isinstance(data, dict):
    
        return {key: convert_encoding(value, from_encoding, to_encoding) for key, value in data.items()}
    elif isinstance(data, list):
        
        return [convert_encoding(item, from_encoding, to_encoding) for item in data]
    return data

# Выполнение SQL-запроса к MS SQL Server
def ms_sql_read(server, database, ad_user, ad_password, querry):
    
    try:
       
        connection = pymssql.connect(server, ad_user, ad_password, database)
        cursor = connection.cursor(as_dict=True)

        cursor.execute(querry)
        result = cursor.fetchall()
        print(result)

        
        cursor.close()
        connection.close()

        result = convert_encoding(result, 'latin1', 'windows-1251')

        return result

    except pymssql.Error as e:
        print(f"An error occurred: {e}")
        return e

# Чтение записи из Redis по ключу userid
def redis_read(userid):
    con = redis.Redis(host=os.getenv('REDIS_URI'), port=6379, db=0)
    try:
        result = json.loads(con.get(userid))# Получение и декодирование JSON
    except TypeError:
        result = None
    finally:
        con.close()
        con.quit()

    print('Redis read result', result, sep='\n')
    return result

# Вставка записи в Redis с TTL (временем жизни)
def redis_insert(data):
    ttl = 3 * 60 # Время жизни ключа в секундах (3 минуты)
    con = redis.Redis(host=os.getenv('REDIS_URI'), port=6379, db=0)
    try:
        con.setex(data['userid'], ttl, json.dumps(data))
        print(f'Successfully inserted data for userid: {data["userid"]}')
    except redis.exceptions.RedisError as e:
        print(f"A redis error occurred: {e}")
    finally:
        result = json.loads(con.get(data['userid']))
        con.close()
        con.quit()

    print(result)
    return result

# Получение информации по заявке из базы данных по номеру или ID
def im_get_call_sql(callnumber=None, callid=None):
    server = os.getenv("SQL_SERVER")
    database = os.getenv("SQL_DATABASE")
    ad_user = os.getenv("AD_USER_NAME")
    ad_password = os.getenv("AD_USER_PASSWORD")
    # Формируем SQL-запрос в зависимости от параметров
    if callid is None:
        querry = (f'SELECT TOP 1 c.*, u.Phone, u.Email, u.PositionName, u.DivisionName, n.Note, n.UserName, n.UtcDate '
                  f'FROM view_Call c '
                  f'FULL JOIN view_User u ON c.ClientID = u.ID '
                  f'FULL JOIN view_CallNote n ON n.CallID = c.ID '
                  f'WHERE c.Number = {callnumber} '
                  f'ORDER BY n.UtcDate DESC')
    else:
        querry = (f"SELECT TOP 1 c.*, u.Phone, u.Email, u.PositionName, u.DivisionName, n.Note, n.UserName, n.UtcDate "
                  f"FROM view_Call c "
                  f"FULL JOIN view_User u ON c.ClientID = u.ID "
                  f"FULL JOIN view_CallNote n ON n.CallID = c.ID "
                  f"WHERE c.ID = '{callid}' "
                  f"ORDER BY n.UtcDate DESC")

    print(querry)

    print(server, database, ad_user, ad_password)

    result = ms_sql_read(server, database, ad_user, ad_password,
                         querry) or f"{callnumber}: Заявка с таким номером не найдена!"

    print(result)

    return result if (isinstance(result, str)) else result[0]

# Получение всех комментариев по заявке
def im_get_call_notes(callid):
    server = os.getenv("SQL_SERVER")
    database = os.getenv("SQL_DATABASE")
    ad_user = os.getenv("AD_USER_NAME")
    ad_password = os.getenv("AD_USER_PASSWORD")
# SQL-запрос для выборки всех комментариев к заявке
    querry = f"SELECT n.*, c.Number FROM view_CallNote n FULL JOIN view_Call c ON n.CallID = c.ID \
               WHERE n.CallID = '{callid}' \
               ORDER BY n.UtcDate ASC"

    result = ms_sql_read(server, database, ad_user, ad_password, querry)

    return result

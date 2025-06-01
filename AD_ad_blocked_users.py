

# Импортируем функцию для получения переменных окружения

from os import getenv

from ldap3 import Server, Connection, NTLM, MODIFY_REPLACE, ALL

# Задаём путь (OU) в Active Directory, в котором будем искать пользователей
OU = 'ou=tass_users,dc=corp,dc=tass,dc=ru'


def ldap_auth():
     #Подключение к серверу Active Directory с использованием учётных данных, указанных в переменных окружения
    server = Server('corp.tass.ru', get_info=ALL)
    conn = Connection(server, user=getenv("AD_USER_NAME"), password=getenv("AD_USER_PASSWORD"), authentication=NTLM)
    conn.bind()

    return conn


def get_ad_blocked_user():
    conn = ldap_auth()
    print(conn)

    # Получение списка заблокированных пользователей из AD
    search_query = '(&(objectClass=user)(lockoutTime>=1)(!(pwdLastSet=0)))'
    # Выполняем поиск по OU
    conn.search(search_base=OU, search_filter=search_query, search_scope='SUBTREE',
                attributes=['name', 'userAccountControl', 'badPwdCount', 'lockoutTime'])
     
    result = []# Список имён заблокированных пользователей
    for entry in conn.entries:
         # Исключаем записи, в которых в пути есть "УВОЛ" (например, уволенные сотрудники),проверяем, что учётная запись активна (userAccountControl == 512)
        if 'УВОЛ' not in entry.entry_dn and entry['userAccountControl'] == 512:
            
            result.append(str(entry['name']))# Добавляем имя пользователя в список

    conn.unbind()# Завершаем соединение с AD

    return sorted(result)# Возвращаем отсортированный список имён



def ad_unblock_user(name):
    # Разблокировка конкретного пользователя по имени. Сбрасывает поле lockoutTime в 0, тем самым снимая блокировку.Возвращает True, если пользователь был разблокирован, иначе False.
    conn = ldap_auth()
    
    search_filter = '(&(objectClass=user)(name=' + name + '))'
    conn.search(OU, search_filter, attributes=['lockoutTime'])

    # Проверяем, заблокирован ли пользователь (lockoutTime ≠ 0)
    if conn.entries[0]['lockoutTime'].value != 0:
        conn.modify(conn.entries[0].entry_dn, {'lockoutTime': [(MODIFY_REPLACE, [0])]})
        conn.unbind()# Завершаем соединение
        return True# Сообщаем, что разблокировка прошла успешно
    else:
        # Если пользователь не был заблокирован — ничего не делаем
        return False

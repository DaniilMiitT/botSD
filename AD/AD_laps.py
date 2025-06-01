from os import getenv

from dotenv import load_dotenv, find_dotenv
from ldap3 import Server, Connection, SUBTREE, ALL_ATTRIBUTES

load_dotenv(find_dotenv('.env.prod'))


def get_laps_password(pc_name):
    # Параметры подключения к LDAP (AD)
    ldap_server = 'corp.tass.ru'
    ldap_username = getenv("AD_USER_NAME")
    ldap_password = getenv("AD_USER_PASSWORD")
    ldap_search_base = 'OU=TASS_Computers,DC=corp,DC=tass,DC=ru'

    # Формируем фильтр для поиска по имени ПК
    pc_filter = '(name={})'.format(pc_name)
    # Указываем, что нас интересует только атрибут с паролем (LAPS)
    pc_attributes = ['ms-Mcs-AdmPwd']

    # Создаём соединение с сервером LDAP и подключаемся автоматически (auto_bind=True)
    server = Server(ldap_server, get_info=ALL_ATTRIBUTES)
    conn = Connection(server, ldap_username, ldap_password, auto_bind=True)

    #Выполняем поиск записи о компьютере по имени
    conn.search(ldap_search_base, pc_filter, SUBTREE, attributes=pc_attributes)
    if conn.entries:
        pc_entry = conn.entries[0]
        if 'ms-Mcs-AdmPwd' in pc_entry:
            laps_password = pc_entry['ms-Mcs-AdmPwd'].value
            if laps_password:
                return f"Пароль LAPS для {pc_name}\r\n```" + laps_password + "```"
            else:
                return f"Пароль LAPS для ПК {pc_name} не найден!"
        else:
            return f"Атрибут с паролем LAPS для ПК {pc_name} не найден в AD!"
    else:
        return f"ПК {pc_name} не найден!"

  
    conn.unbind()

    return None

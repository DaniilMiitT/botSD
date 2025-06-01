from os import getenv

from dotenv import load_dotenv, find_dotenv
from ldap3 import Server, Connection, ALL, NTLM

load_dotenv(find_dotenv('.env.prod'))


def auth_vkteams_user(email, group):
#Проверяет, состоит ли пользователь с указанной почтой `email` в нужной AD-группе `group`.Используется для авторизации пользователей в боте.Возвращает True, если пользователь найден в нужной группе, иначе — False.
    
    server = Server('corp.tass.ru', get_info=ALL)
    # Получаем логин и пароль сервисного аккаунта из переменных окружения
    login = getenv("AD_USER_NAME")
    password = getenv("AD_USER_PASSWORD")
    # Область поиска в AD
    ou = "OU=TASS_Users,DC=corp,DC=tass,DC=ru"
    allowed_group = f"CN={group},OU=VKTeams_Bot_Users,OU=VKTeams,OU=Groups,DC=corp,DC=tass,DC=ru"
    attribs = ['mail']

    conn = Connection(server, user=login, password=password, authentication=NTLM, auto_bind=True)
    conn.search(search_base=ou, search_filter="(memberOf=" + allowed_group + ")", attributes=attribs)

    try:
        if conn.entries:
            print(conn.entries)
            for i in conn.entries:
                if email in i['mail']:
                    return True
        return False
    except ValueError:
        return False

# auth_vkteams_user("ponomarev_d@tass.ru", "SD_Tier3")

import requests
from packaging import version
from requests_ntlm import HttpNtlmAuth
from os import getenv


def sccm_get_app_versions():
    #Получает список всех приложений из SCCM через AdminService API.
    #Возвращает словарь, где ключ — ID приложения, а значение — его имя и версия.
    url = 'https://msk-sccm-ss02.corp.tass.ru/AdminService/wmi/SMS_Application'

    #Создаем сессию с NTLM-аутентификацией
    session = requests.Session()
    session.auth = HttpNtlmAuth(getenv("AD_USER_NAME"), getenv("AD_USER_PASSWORD"))
    # Отправляем GET-запрос к SCCM API
    response = session.get(url, verify=False).json()['value']

    result = {}
# Формируем словарь с данными о приложениях
    for i in response:
        result[i["CI_ID"]] = {
            'name': i['LocalizedDisplayName'], # Локализованное имя приложения
            'version': i['SoftwareVersion'] # Версия
        }

    return result


def sccm_get_latest_app_version(apps, app_name):
    # Из полученного словаря приложений возвращает последнюю (наивысшую) версию указанного приложения.
   # Параметры:
      #  apps: словарь приложений, возвращенный sccm_get_app_versions()
       # app_name: имя приложения для поиска
    #Возвращает:
      #  {'Name': ..., 'Version': ...} — если найдено;
        #пустой словарь — если не найдено.
      
    result = {}
    max_version = version.parse("0.0.0.0")# Начальное значение для сравнения
    for app_id, app_info in apps.items():
         # Ищем совпадение по имени (без учета регистра) и обновляем максимум, если версия новее
        if app_name.lower() in app_info['name'].lower() and version.parse(app_info['version']) > max_version:
            
            max_version = version.parse(app_info['version'])
            )
            result = {'Name': app_info['name'], 'Version': max_version}

    return result

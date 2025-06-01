import os

import requests
from fastapi import APIRouter
from packaging import version
from requests_ntlm import HttpNtlmAuth

sccm_app_versions = APIRouter()


@sccm_app_versions.get("/sccm_app_versions/{appname}")
def get_sccm_app_versions(appname: str):
    # Формируем URL запроса к SCCM API с фильтром по имени приложения, начинающемуся с appname
    url = f"https://msk-sccm-ss02.corp.tass.ru/AdminService/wmi/SMS_Application?$filter=startswith(LocalizedDisplayName,%27{appname}%27)%20eq%20true"
    
    session = requests.Session()

     # Авторизация через NTLM с помощью переменных окружения для пользователя и пароля AD
    session.auth = HttpNtlmAuth(os.getenv("AD_USER_NAME"), os.getenv("AD_USER_PASSWORD"))

    # Выполняем GET запрос к SCCM API
    response = session.get(url, verify=False).json()['value']
    

    result = {}
    latest_version = version.parse('0.0.0.0')# Инициализируем "минимальную" версию

    for item in response:
        # Выводим для отладки имя и версию приложения
        print(item.get('LocalizedDisplayName'), item.get('SoftwareVersion'))
        try:
            sofware_version = version.parse(item.get('SoftwareVersion'))
        except TypeError as e:
             # Если версия отсутствует или некорректна — возвращаем ошибку
            return {'Error in getting software version': str(e)}

        # Сравниваем текущую версию с найденной максимальной
        if sofware_version > latest_version:
            latest_version = version.parse(item.get("SoftwareVersion"))
            result = {
                'Name': item.get('LocalizedDisplayName'),
                'Version': item.get("SoftwareVersion")
            }
            # Если ничего не найдено, возвращаем соответствующее сообщение
    return result or {'Result': f'Nothing found in SCCM for {appname}'}

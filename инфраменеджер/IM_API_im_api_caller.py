import os

import requests
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv('.env.prod'))


# Получаем основной URI апи инфраменеджер из переменных окружения
main_uri = os.getenv("IM_API_URI")


def im_auth():
    # Получаем логин и пароль из переменных окружения
    login = os.getenv("IM_API_USER")
    password = os.getenv("IM_API_PASSWORD")

    key = 13
    passwordEncrypted = ''
    for i in range(len(password)):
        passwordEncrypted += chr(ord(password[i]) ^ key)
# URI для авторизации
    api_uri = "/accountApi/SignIn"
    session = requests.Session()

    body = {
        "loginName": login,
        "passwordEncrypted": passwordEncrypted,
    }
# Выполняем POST-запрос на авторизацию
    try:
        response = session.post(url=main_uri + api_uri, data=body, verify=False)
        print("Auth success: ", response.json()['Success'])
    except Exception as e:
        print('Ошибка', str(e))

    return session# Возвращаем сессию с авторизацией


def im_search_call_by_executor(executor):
     # Ищем заявки, в которых указан исполнитель
    session = im_auth()
    uri = main_uri + "/sdApi/GetListForObject"
    body = {
        'StartRecordIndex': 0,
        'CountRecords': 160000,
        'ViewName': 'CallForTable',
        'TimezoneOffsetInMinutes': 0,
        'FirstTimeLoad': 'false',
    }
# Запрашиваем список заявок
    response = session.post(url=uri, data=body, verify=False)
    result = []
    # Фильтруем заявки по имени исполнителя
    for i in response.json()['Data']:
        if executor in i['ExecutorFullName']:
            result.append(i)

    return result


def im_get_user_by_mail(userid):
    # Получаем пользователя по email
    session = im_auth()
    uri = main_uri + "/searchApi/search"

    body = {
        'Text': userid,
        'TypeName': 'ExecutorUserSearcher',
        'Params': '[]',
    }

    result = session.post(url=uri, data=body, verify=False)

    return result.json()[0]


def im_get_user_by_id(userid):
    # Получаем полную информацию о пользователе по его ID
    session = im_auth()
    uri = main_uri + f"/userApi/GetUserInfo?userID={userid}"

    result = session.get(url=uri, verify=False)

    return result.json()


def im_set_call_field_value(callid, field, value):
    # Устанавливаем значение для поля в заявке
    session = im_auth()
    uri = main_uri + "/sdApi/SetField"

    body = {
        'ID': callid,
        'ObjClassID': 701,# Класс объекта "Заявка"
        'Field': field,
        'OldValue': '',
        'NewValue': '',
        'ReplaceAnyway': True,
    }
    #Если поле — решение, форматируем в виде текста
    if field == 'Call.Solution':
        body['OldValue'] = "{'text': ''}"
        body['NewValue'] = "{'text': '" + value + "'}"
        
    elif field == 'Call.Executor' or field == 'Call.Accomplisher':
        body['OldValue'] = "{'id':'','fullName':'','classID':9}"
        body['NewValue'] = "{'id':'value1','fullName':'value2','classID':9}".replace(
            'value1', value['ID']).replace('value2', value['FullName'])
# Отправка запроса на изменение поля
    response = session.post(url=uri, data=body, verify=False)
# Возвращаем результат выполнения
    return f"Set {field} success" if response.json()['ResultWithMessage']['Result'] == 0 \
        else response.json()['ResultWithMessage']['Message']


def im_add_note(callid, text, type):
    # Добавляем комментарий или сообщение в заявку
    session = im_auth()
    api_uri = main_uri + "/sdApi/AddNote"
 
    body = {
        "Id": callid,
        "entityClassID": 701,
        "Message": text,
        "Type": type,  # 1 = msg, 0 = note
    }

    response = session.post(url=api_uri, data=body, verify=False)

    return response# Возвращаем ответ от API


def im_set_call_state(stateid, callid):
    # Удаляем заявку по ID
    session = im_auth()
    api_uri = main_uri + f"/workflowApi/setState?entityID={callid}&entityClassID=701&targetStateID={stateid}"

    response = session.post(url=api_uri, verify=False)

    return response.json()# Возвращаем ответ от сервера


def im_remove_object(call_id):
    session = im_auth()
    api_uri = main_uri + '/sdApi/RemoveObjectList'
    body = {
        "ObjectList[0][ID]": call_id,
        "ObjectList[0][ClassID]": "701",
    }
    response = session.post(url=api_uri, data=body, verify=False)

    return response.json()['Result']

